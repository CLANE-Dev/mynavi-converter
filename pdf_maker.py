"""LibreOffice headless で xlsx を PDF に変換する。"""
from __future__ import annotations

import glob
import logging
import os
import shutil
import subprocess
import tempfile

import config

log = logging.getLogger(__name__)

SOFFICE = shutil.which("soffice") or "/usr/bin/soffice"


def _run(args: list[str], timeout: int, profile: str) -> subprocess.CompletedProcess:
    env = dict(os.environ)
    env.setdefault("HOME", "/tmp")
    return subprocess.run(
        args + [f"-env:UserInstallation=file://{profile}"],
        capture_output=True,
        timeout=timeout,
        env=env,
        check=False,
    )


def warmup() -> None:
    """初回変換が遅くならないよう、起動時に LibreOffice を一度回す。"""
    try:
        import openpyxl

        with tempfile.TemporaryDirectory() as tmp:
            src = os.path.join(tmp, "warmup.xlsx")
            wb = openpyxl.Workbook()
            wb.active["A1"] = "ウォームアップ"
            wb.save(src)
            to_pdf(src, tmp)
        log.info("LibreOffice ウォームアップ完了")
    except Exception as exc:
        log.warning("ウォームアップ失敗（無視して継続）: %s", exc)


def to_pdf(xlsx_path: str, out_dir: str) -> str:
    os.makedirs(out_dir, exist_ok=True)
    with tempfile.TemporaryDirectory() as profile:
        proc = _run(
            [
                SOFFICE,
                "--headless",
                "--norestore",
                "--nolockcheck",
                "--nodefault",
                "--nofirststartwizard",
                "--convert-to",
                "pdf:calc_pdf_Export",
                "--outdir",
                out_dir,
                xlsx_path,
            ],
            timeout=config.SOFFICE_TIMEOUT,
            profile=profile,
        )

    stem = os.path.splitext(os.path.basename(xlsx_path))[0]
    expected = os.path.join(out_dir, stem + ".pdf")
    if os.path.exists(expected) and os.path.getsize(expected) > 0:
        return expected

    found = sorted(glob.glob(os.path.join(out_dir, "*.pdf")))
    if found:
        return found[0]

    raise RuntimeError(
        "PDF変換に失敗しました rc=%s stderr=%s"
        % (proc.returncode, proc.stderr.decode("utf-8", "ignore")[:400])
    )
