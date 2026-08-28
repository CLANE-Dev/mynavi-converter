"""マイナビ請求書 変換サービス。

役割は1つだけ。パスワード付き xlsx を受け取り、自社情報と角印を入れて
PDF にして返す。Google にも Slack にもアクセスしない。受け取ったファイルは
応答後に削除し、ディスクにもログにも中身を残さない。
"""
from __future__ import annotations

import base64
import hmac
import logging
import os
import tempfile

from fastapi import FastAPI, Header, HTTPException
from fastapi.responses import JSONResponse
from pydantic import BaseModel, Field

import config
import excel_editor
import pdf_maker

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(name)s %(message)s",
)
log = logging.getLogger("mynavi-converter")

VERSION = "1.1"
PREVIEW_WIDTH_PX = int(os.environ.get("PREVIEW_WIDTH_PX", "1100"))
app = FastAPI(title="mynavi-converter", version=VERSION, docs_url=None, redoc_url=None)


class ConvertRequest(BaseModel):
    filename: str = Field(..., max_length=300)
    xlsx_b64: str
    kind: str = Field("auto", pattern="^(auto|invoice|delivery)$")
    preview: bool = False
    stamp_b64: str | None = None


@app.on_event("startup")
def _startup() -> None:
    if not config.SHARED_SECRET:
        log.error("SHARED_SECRET が未設定です。/convert は常に401を返します")
    if not config.EXCEL_PASSWORD:
        log.error("EXCEL_PASSWORD が未設定です。復号に失敗します")
    pdf_maker.warmup()


@app.get("/health")
def health() -> dict:
    return {
        "ok": True,
        "service": "mynavi-converter",
        "version": VERSION,
        "secret_configured": bool(config.SHARED_SECRET),
        "password_configured": bool(config.EXCEL_PASSWORD),
    }


def _authorize(provided: str | None) -> None:
    if not config.SHARED_SECRET:
        raise HTTPException(status_code=401, detail="サーバー側の合言葉が未設定です")
    if not provided or not hmac.compare_digest(provided, config.SHARED_SECRET):
        raise HTTPException(status_code=401, detail="合言葉が一致しません")


def _is_invoice(filename: str, kind: str) -> bool:
    if kind == "invoice":
        return True
    if kind == "delivery":
        return False
    # auto: ファイル名で判定。「納品」を含めば納品書、それ以外は請求書扱い。
    return "納品" not in filename


@app.post("/convert")
def convert(
    req: ConvertRequest,
    x_shared_secret: str | None = Header(default=None, alias="X-Shared-Secret"),
):
    _authorize(x_shared_secret)

    try:
        raw = base64.b64decode(req.xlsx_b64, validate=True)
    except Exception:
        raise HTTPException(status_code=400, detail="xlsx_b64 をデコードできません")
    if not raw:
        raise HTTPException(status_code=400, detail="xlsx が空です")
    if len(raw) > config.MAX_UPLOAD_BYTES:
        raise HTTPException(status_code=413, detail="xlsx が大きすぎます")

    stamp = None
    if req.stamp_b64:
        try:
            stamp = base64.b64decode(req.stamp_b64, validate=True)
        except Exception:
            raise HTTPException(status_code=400, detail="stamp_b64 をデコードできません")
        if len(stamp) > config.MAX_UPLOAD_BYTES:
            raise HTTPException(status_code=413, detail="角印画像が大きすぎます")

    base = os.path.basename(req.filename) or "document.xlsx"
    stem = os.path.splitext(base)[0]
    is_invoice = _is_invoice(base, req.kind)

    with tempfile.TemporaryDirectory(prefix="mynavi-") as work:
        edited = os.path.join(work, "edited", stem + ".xlsx")
        try:
            result = excel_editor.edit(
                raw,
                password=config.EXCEL_PASSWORD,
                is_invoice=is_invoice,
                stamp_png=stamp,
                out_path=edited,
            )
        except excel_editor.TemplateChanged as exc:
            log.warning("テンプレート差分により中断: %s", exc)
            return JSONResponse(
                status_code=409,
                content={
                    "ok": False,
                    "error": "template_changed",
                    "detail": str(exc),
                },
            )
        except Exception as exc:
            log.exception("記入処理に失敗")
            return JSONResponse(
                status_code=500,
                content={"ok": False, "error": "edit_failed", "detail": str(exc)[:300]},
            )

        try:
            pdf_path = pdf_maker.to_pdf(edited, os.path.join(work, "pdf"))
        except Exception as exc:
            log.exception("PDF変換に失敗")
            return JSONResponse(
                status_code=500,
                content={"ok": False, "error": "pdf_failed", "detail": str(exc)[:300]},
            )

        with open(pdf_path, "rb") as fh:
            pdf_bytes = fh.read()

    log.info(
        "変換完了 %s invoice=%s stamped=%s pdf=%dbytes",
        base,
        is_invoice,
        result["stamped"],
        len(pdf_bytes),
    )
    return {
        "ok": True,
        "filename": stem + ".pdf",
        "is_invoice": is_invoice,
        "stamped": result["stamped"],
        "sheet": result["sheet"],
        "pdf_size": len(pdf_bytes),
        "pdf_b64": base64.b64encode(pdf_bytes).decode("ascii"),
    }
