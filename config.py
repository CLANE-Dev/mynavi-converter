"""固定値と記入セルの定義。

セル位置は 2026年7月の実テンプレートで検証済み。
テンプレートのセルには最初から「住所：」等のラベルが入っているため、
値で上書きせずラベルの後ろへ追記する（excel_editor._write 参照）。
"""
import json
import os

INFO = {
    "zip": os.environ.get("COMPANY_ZIP", "〒102-0081"),
    "addr": os.environ.get(
        "COMPANY_ADDR", "東京都千代田区四番町2番地1-1 クレール東郷坂 1F"
    ),
    "tel": os.environ.get("COMPANY_TEL", "08033861900"),
    "reg": os.environ.get("COMPANY_REG", "T4012301011308"),
    "bank": os.environ.get("BANK_NAME", "三井住友銀行"),
    "branch": os.environ.get("BANK_BRANCH", "麹町支店"),
    "acc_no": os.environ.get("BANK_ACC_NO", "9406859"),
    "acc_name": os.environ.get("BANK_ACC_NAME", "カ）クライン"),
    "note": os.environ.get("INVOICE_NOTE", ""),
}

TARGET_SHEET = "表紙"

COMMON_CELLS = {
    "F13": INFO["zip"],
    "E14": INFO["addr"],
    "E20": INFO["tel"],
}

INVOICE_CELLS = {
    "E22": INFO["reg"],
    "B36": INFO["bank"],
    "B37": INFO["branch"],
    "B39": INFO["acc_no"],
    "B41": INFO["acc_name"],
    "B45": INFO["note"],
}

# F13 は元から空欄の入力枠なのでガードしない。
# 環境変数 GUARD_LABELS_JSON / GUARD_LABELS_INVOICE_JSON で上書きできる。
_DEFAULT_LABELS = {
    "E14": ("住所",),
    "E20": ("TEL", "ＴＥＬ", "電話"),
}
_DEFAULT_LABELS_INVOICE = {
    "E22": ("登録", "番号"),
    "B36": ("銀行",),
    "B37": ("支店",),
    "B39": ("口座",),
    "B41": ("名義",),
}


def _labels_from_env(var, fallback):
    raw = os.environ.get(var, "").strip()
    if not raw:
        return fallback
    try:
        parsed = json.loads(raw)
        return {k: tuple(v) for k, v in parsed.items()}
    except Exception:
        return fallback


EXPECTED_LABELS = _labels_from_env("GUARD_LABELS_JSON", _DEFAULT_LABELS)
EXPECTED_LABELS_INVOICE = _labels_from_env(
    "GUARD_LABELS_INVOICE_JSON", _DEFAULT_LABELS_INVOICE
)

STAMP_ANCHOR = os.environ.get("STAMP_ANCHOR", "H18")
STAMP_WIDTH_PX = int(os.environ.get("STAMP_WIDTH_PX", "105"))
STAMP_OFFSET_X_PX = int(os.environ.get("STAMP_OFFSET_X_PX", "0"))
STAMP_OFFSET_Y_PX = int(os.environ.get("STAMP_OFFSET_Y_PX", "0"))

EXCEL_PASSWORD = os.environ.get("EXCEL_PASSWORD", "")
SHARED_SECRET = os.environ.get("SHARED_SECRET", "")
SOFFICE_TIMEOUT = int(os.environ.get("SOFFICE_TIMEOUT", "120"))
MAX_UPLOAD_BYTES = int(os.environ.get("MAX_UPLOAD_BYTES", str(8 * 1024 * 1024)))
