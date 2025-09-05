# デモ機予約管理Bot 設計書

**設計書バージョン:** 5.1  
**バージョン名:** Chatbotステートマシン実装版  
**更新日時:** 2025.09.05

## 1. システムアーキテクチャ

### 1.1 概要
ローカル PC 上で完結する Web アプリケーション。  
フロントエンド（ブラウザ）は Flask サーバーの Chat API に HTTP/JSON でアクセスし、  
サーバーは Excel ファイルをデータストアとして読み書きします。  
保存方式は既定で openpyxl、互換性保持が必要な場合は Excel COM（win32com）を使用します。

### 1.2 コンポーネント構成
```mermaid
flowchart LR
  U[ユーザー] --> B[Webブラウザ]
  B -- HTTP/JSON --> S[Flask サーバー<br/>Chat API]
  S -- 読み書き --> X[(Excel ファイル<br/>デモ機予約表.xlsx)]
  S <-. オプション:COM .-> E[Excel(アプリケーション)]

  subgraph クライアント
    B
  end
  subgraph サーバー
    S
  end
  subgraph データストア
    X
  end
```

- Webブラウザ (frontend): `project/src/static/` 内の HTML・JS。チャット UI と `sessionStorage` による状態保持。  
- Flask サーバー (backend): `project/src/app.py`。ステートマシンで対話を制御し `excel_ops` を呼び出す。  
- Excel 操作モジュール: `project/src/excel_ops.py`。在庫確認・予約・キャンセル・一覧取得。openpyxl または COM 経由の I/O。  
- Excel ファイル: `project/data/デモ機予約表.xlsx`。月別シートと予約ログを保持。  

### 1.3 データフロー
1. ブラウザが `POST /api/chat` に `state`・`user_info`・`context` と入力テキストを送信。  
2. Flask がステートマシンで意図を判定（予約／キャンセル／確認など）。  
3. `excel_ops` が Excel を読み書き。  
   - 予約: 空き確認 → セルに `C:<予約ID>` 記入 → 予約ログへ追記（シート自動作成・黒字化）。  
   - キャンセル: 予約IDで期間セルを特定しクリア。ログのステータスを更新。  
   - 確認: 予約ログからユーザーの予約一覧を抽出。  
4. レスポンスとして応答文・次状態・更新済み `user_info`/`context` を返却。  

### 1.4 実行モード（保存方式）
- openpyxl (既定)  
  - ディレクトリロック、Excel 開放待ち、テンポラリ保存→`os.replace`、`.bak` 退避。  
  - ZIP 構造・シート存在・サイズ差 (`EXCEL_SIZE_DIFF_RATIO`, 既定 0.5) を検証し異常時はロールバック。  
  - 条件付き書式・データ検証・図形など一部機能は失われる可能性。  
- COM モード  
  - 環境変数 `EXCEL_WRITE_MODE=com` で有効化（Windows + Excel + pywin32 必須）。  
  - `pythoncom.CoInitialize/CoUninitialize` を各操作で呼出し、Excel 本体で編集・保存。  

### 1.5 環境変数
| 変数名 | 意味 | 既定値 |
| :--- | :--- | :--- |
| `EXCEL_WRITE_MODE` | `com` で COM モード、それ以外は openpyxl | (空) |
| `EXCEL_SIZE_DIFF_RATIO` | 保存前後のファイルサイズ差閾値 (0.0–1.0) | `0.5` |

### 1.6 主要ステート (サーバー)
- `AWAITING_USER_INFO_*` : 名前 → 内線 → 職番 の順で取得  
- `AWAITING_COMMAND` : コマンド待ち（予約 / キャンセル / 確認）  
- 予約系 : `AWAITING_DEVICE_TYPE` → `AWAITING_DATES` → `CONFIRM_RESERVATION`  
- キャンセル系 : `AWAITING_CANCEL_BOOKING_ID` → `CANCEL_CONFIRM`  
- 確認系 : コマンド後すぐ一覧を表示し `AWAITING_COMMAND` へ戻る  

---

## 2. データモデル設計 (Excel)

### 2.1 ファイル構成
- 位置: `project/data/デモ機予約表.xlsx`  
- 構成: 月別カレンダーシート（複数） + `予約ログ` シート  

### 2.2 カレンダーシート（`yy年M月`）
- シート名例: `25年9月`, `25年10月`  
- ヘッダー行: 8 行目の C8 から右方向に当月 1..末日を配置  
  - 許容フォーマット: `1`, `1日`, 全角数字など（内部で正規化して解釈）  
- デモ機名: B 列（B9 から下方向）  
- 予約データ領域: C9 以降の格子（行=デモ機, 列=日）  
- セル値の意味:  
  - 空欄 … 空き  
  - `C:<予約ID>` … 指定 ID の予約  

### 2.3 予約ログシート（`予約ログ`）
- シートは存在しない場合、自動で作成  
- 列（推奨順）:  
  - 予約ID（文字列）  
  - 予約日時（JST, `yyyy-MM-ddTHH:mm:ss`）  
  - 予約者名（文字列）  
  - 内線番号（文字列）  
  - 職番（文字列）  
  - デモ機名（文字列）  
  - 予約開始日（`YYYY-MM-DD`）  
  - 予約終了日（`YYYY-MM-DD`）  
  - ステータス（`予約中` / `キャンセル済` など）  
- 可読性: 追記時にフォント色は黒に統一  

### 2.4 月またぎ予約の仕様
1. 同一デモ機を全期間で確保（途中で機種変更しない）  
2. 予約IDは 1 つ（ログは 1 行）  
3. いずれかの月シートが無い場合はエラーとして予約を成立させない  
4. 実装方針: [開始, 終了] を月ごとに分割し、各サブ区間で空き確認→書込み→（キャンセル時は）クリア  
5. 予約ログには実際の開始日・終了日を保存  
6. 利用関数: `find_available_device`, `book`, `cancel` で共通化  

---

## 3. API設計 (ローカルサーバー)

### 3.1 Chatbot 対話 API
- エンドポイント: `POST /api/chat`  
- リクエスト (JSON):  
  ```json
  {
    "text": "予約",                 
    "state": "AWAITING_COMMAND",  
    "user_info": {
      "name": "山田太郎",
      "extension": "1234",
      "employee_id": "A001"
    },
    "context": {}
  }
  ```
- レスポンス (JSON):  
  ```json
  {
    "reply_text": "ご希望のデモ機の種類を入力してください。",
    "next_state": "AWAITING_DEVICE_TYPE",
    "user_info": {"name": "山田太郎", "extension": "1234", "employee_id": "A001"},
    "context": {"intent": "reserve"}
  }
  ```
- ステータスコード:  
  - 200: 成功  
  - 400: JSON 形式エラー  
  - 500: サーバー内部エラー（原則ログに記録しユーザーには簡潔に通知）  

レスポンス例（予約一覧表示）  
```
あなたの予約一覧:
- 19fba056 [予約中] FE-01 2025-09-10→2025-09-12
- 98c4a221 [キャンセル済] PC-03 2025-08-28→2025-08-30
```

### 3.2 内部ロジック（要点）
- 発話理解は定型コマンドで判定: 「予約」/「キャンセル」/「確認|予約確認|予約状況」  
- ステートマシンで遷移を制御し、必要に応じて `excel_ops` を呼び出す  
  - 予約: `find_available_device` → `book`  
  - キャンセル: `list_cancellable_bookings` → `cancel`  
  - 確認: `list_user_bookings`  
- Excel 保存は `excel_ops` で一元管理（ファイルロック、原子的保存、.bak 作成、検証、COM モード）  

### 3.3 予約確認（自分の予約一覧）
- トリガー語: 「確認」 / 「予約確認」 / 「予約状況」  
- 動作: 予約ログから、現在のユーザー（名前・内線・職番のいずれか一致）に紐づく予約を最大 10 件まで一覧表示  
- 表示形式:  
  `- <予約ID> [<ステータス>] <デモ機名> <開始日>→<終了日>`  
- 備考:  
  - 一致条件は OR（名前一致、内線一致、職番一致のいずれか）  
  - 「予約中」以外のステータス（キャンセル済など）も表示対象  
  - 一覧は読み取り専用。キャンセルは「キャンセル」コマンドで実行  

---

## 4. フロントエンド設計
- UI: `index.html` + `static/main.js` によるシンプルなチャット画面  
- セッション管理:  
  - ページ読み込み時に `sessionStorage` をリセットし常に新規会話を開始  
  - `state`, `user_info`, `context` を保持し、各入力ごとに `POST /api/chat` に送信  
  - API 応答で前記 3 つを更新し、レンダリング後はオートスクロール  
- 初回メッセージ: ページロード直後に自動送信し、名前入力から開始  
- エラー表示: fetch 失敗時に System メッセージを表示  

---

## 5. 開発/運用
- 言語: Python 3.13（venv 同梱）  
- 主要ライブラリ: Flask, openpyxl, python-dotenv, pytest, （COM 時）pywin32  
- テスト: `python -m pytest -q` → 22 passed, 2 xfailed, 1 xpassed  
- 起動:  
  ```powershell
  # 通常
  python src/app.py
  # COM モード
  $env:EXCEL_WRITE_MODE = "com"
  python src/app.py
  ```
- 運用注意: Excel ファイルを Excel で開いたままにしない（サーバーは開放待ち、タイムアウトでエラー）  
