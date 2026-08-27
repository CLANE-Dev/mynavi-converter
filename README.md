# mynavi-converter

マイナビ月次請求書の自動処理のうち、Apps Script では実行できない
「パスワード付き xlsx の復号 → 自社情報の記入 → 角印 → PDF化」だけを担当する小さなサービス。

## 全体の流れ

```
メール着信 (mj.keiri@mynavi.jp)
  → GAS checkAndProcess      1分毎  xlsx を Drive 26年/{月}/エクセル/ に保存
  → GAS convertAndStage_     1分毎  xlsx と角印をこのサービスへ送る
        → mynavi-converter          復号・記入・角印・PDF化して返す（保存しない）
     ← PDF を 送信前PDF/ に保存 → 続けて下書き生成を呼ぶ
  → GAS createMynaviDraftIfNeeded  添付付き下書き + Slack通知
  → 人が下書きを送信
  → GAS collectSentPdfs      10分毎 送信を検知して PDF/ へ格納
```

着信から下書きまで通常2分以内。

## 設計上の前提

- Google にも Slack にもアクセスしない。認証情報を一切持たない。
- 受け取ったファイルは一時ディレクトリで処理し、応答後に消える。ディスクにもログにも中身を残さない。
- 通信は合言葉（`X-Shared-Secret` ヘッダ）で保護する。一致しなければ401。
- テンプレートの様式が変わったと判断したら 409 を返して変換しない。
  誤った請求書を送るより止まるほうが安全という判断。

## 環境変数

| 変数 | 必須 | 既定 | 内容 |
|---|---|---|---|
| `SHARED_SECRET` | 必須 | なし | GAS 側 `CONVERTER_SECRET` と同じ値 |
| `EXCEL_PASSWORD` | 必須 | なし | マイナビ xlsx の解除パスワード |
| `TZ` | 推奨 | Asia/Tokyo | |
| `COMPANY_TEL` | 任意 | 08033861900 | 請求書に載る電話番号 |
| `COMPANY_ZIP` / `COMPANY_ADDR` / `COMPANY_REG` | 任意 | 実値 | |
| `BANK_NAME` / `BANK_BRANCH` / `BANK_ACC_NO` / `BANK_ACC_NAME` | 任意 | 実値 | |
| `INVOICE_NOTE` | 任意 | 空 | 備考欄（B45）。空なら何も書かない |
| `STAMP_ANCHOR` | 任意 | H18 | 角印を貼るセル |
| `STAMP_WIDTH_PX` | 任意 | 105 | 角印の幅。位置や大きさの微調整はここだけで済む |

## API

### GET /health

```json
{"ok":true,"service":"mynavi-converter","version":"1.0",
 "secret_configured":true,"password_configured":true}
```

### POST /convert

ヘッダ `X-Shared-Secret: <合言葉>`

```json
{
  "filename": "株式会社CLANE様8月25日締め請求書テンプレート_+Digital編集課.xlsx",
  "xlsx_b64": "...",
  "stamp_b64": "...",
  "kind": "auto"
}
```

`kind` は `auto`（ファイル名で判定）/ `invoice` / `delivery`。
ファイル名に「納品」を含めば納品書、それ以外は請求書として扱う。
請求書のときだけ登録番号と振込先を記入する。

応答（200）

```json
{"ok":true,"filename":"....pdf","is_invoice":true,"stamped":true,
 "sheet":"表紙","pdf_size":36906,"pdf_b64":"..."}
```

主なエラー

| コード | 意味 | GAS側の扱い |
|---|---|---|
| 401 | 合言葉不一致 | 失敗通知 |
| 409 | テンプレート様式が想定と違う | 変換せず停止し、テンプレ変更として通知 |
| 500 | 記入またはPDF変換に失敗 | 失敗通知（下書きは作らない） |

## 記入セル（2026年7月の実テンプレートで検証済み）

テンプレートのセルには最初から「住所：」等のラベルが入っているため、
値で上書きせずラベルの後ろへ追記する。二重実行しても重複しない。

| セル | 内容 | 請求書のみ |
|---|---|---|
| F13 | 郵便番号 | |
| E14 | 住所 | |
| E20 | 電話番号 | |
| E22 | 登録番号 | ○ |
| B36 / B37 / B39 / B41 | 銀行名 / 支店名 / 口座番号 / 口座名義 | ○ |
| B45 | 備考 | ○ |

シート構成は 注意書き / 表紙 / 明細 の順で、記入対象は「表紙」。
非表示シートは LibreOffice の PDF 出力に含まれないため、出力は表紙＋明細の2ページになる。

## デプロイ

Dokploy（CLANE ONE の provision_product）で配備する。

- build_type: dockerfile
- port: 8080
- needs_postgres: false
- 公開ドメイン: mynavi-converter.clane.co.jp

## GAS 側の接続

1. `convert_and_stage.gs` の全文を コード.gs の末尾へ貼る（既存コードは変更しない）
2. スクリプトプロパティに `CONVERTER_URL`（`https://.../convert`）と `CONVERTER_SECRET` を追加
3. 関数 `installConverterTrigger` を1回実行

プロパティが未設定のうちは `convertAndStage_` は何もしないので、1を先にやっても害はない。

動作確認は `pingConverter`（/health 疎通）と `runConverterNow`（即時1回実行）で行う。

## 止まったときの見方

- Slack に通知が出ていない かつ 送信前PDF/ が空 → GAS のトリガーを疑う（`installConverterTrigger` 再実行）
- 「接続できません」通知 → Dokploy でコンテナの状態を見る
- 「テンプレートの様式が変わっている可能性」通知 → エクセル/ の原本を開いて対象セルを確認し、`config.py` の該当セルを直す
