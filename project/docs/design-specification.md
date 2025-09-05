# デモ機予約管理Bot 設計書

**設計書バージョン:** 5.0  
**バージョン名:** Chatbotステートマシン実装版  
**更新日時:** 2025.09.03

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

- **Webブラウザ (frontend)**: `static/` 内の HTML・JS。チャット UI と `sessionStorage` による状態保持。  
- **Flask サーバー (backend)**: `src/app.py`。ステートマシンで対話を制御し `excel_ops` を呼び出す。  
- **Excel 操作モジュール**: `src/excel_ops.py`。在庫確認・予約・キャンセル・一覧取得。openpyxl または COM 経由の I/O。  
- **Excel ファイル**: `data/デモ機予約表.xlsx`。月別シートと予約ログを保持。  

### 1.3 データフロー
1. ブラウザが `POST /api/chat` に `state`・`user_info`・`context` と入力テキストを送信。  
2. Flask がステートマシンで意図を判定（予約／キャンセル／確認など）。  
3. `excel_ops` が Excel を読み書き。  
   - **予約**: 空き確認 → セルに `C:<予約ID>` 記入 → 予約ログへ追記（シート自動作成・黒字化）。  
   - **キャンセル**: 予約IDで期間セルを特定しクリア。  
   - **確認**: 予約ログからユーザーの予約一覧を抽出。  
4. レスポンスとして応答文・次状態・更新済み `user_info`/`context` を返却。  

### 1.4 実行モード（保存方式）
- **openpyxl (既定)**  
  - ディレクトリロック、Excel 開放待ち、テンポラリ保存→`os.replace`、`.bak` 退避。  
  - ZIP 構造・シート存在・サイズ差 (`EXCEL_SIZE_DIFF_RATIO`, 既定 0.5) を検証し異常時はロールバック。  
  - 条件付き書式・データ検証・図形など一部機能は失われる可能性。  
- **COM モード**  
  - 環境変数 `EXCEL_WRITE_MODE=com` で有効化（Windows + Excel + pywin32 必須）。  
  - `pythoncom.CoInitialize/CoUninitialize` を各操作で呼出し、Excel 本体で編集・保存。  

### 1.5 環境変数
| 変数名 | 意味 | 既定値 |
| :--- | :--- | :--- |
| `EXCEL_WRITE_MODE` | `com` で COM モード、それ以外は openpyxl | *(空)* |
| `EXCEL_SIZE_DIFF_RATIO` | 保存前後のファイルサイズ差閾値 (0.0–1.0) | `0.5` |

### 1.6 主要ステート (サーバー)
- `AWAITING_USER_INFO_*` : 名前 → 内線 → 職番 の順で取得  
- `AWAITING_COMMAND` : コマンド待ち（予約 / キャンセル / 確認）  
- **予約系** : `AWAITING_DEVICE_TYPE` → `AWAITING_DATES` → `CONFIRM_RESERVATION`  
- **キャンセル系** : `AWAITING_CANCEL_BOOKING_ID` → `CANCEL_CONFIRM`  
- **確認系** : コマンド後すぐ一覧を表示し `AWAITING_COMMAND` へ戻る  

## 2. データモデル設計 (Excel)
*(unchanged)*

## 3. API設計 (ローカルサーバー)

### 3.1. Chatbot対話API  
*(抜粋のみ変更点を記載)*

レスポンス例（予約一覧表示）  
```
あなたの予約一覧:
- 19fba056 [予約中] FE-01 2025-09-10→2025-09-12
- 98c4a221 [キャンセル済] PC-03 2025-08-28→2025-08-30
```

### 3.2. 内部ロジック
*(既存内容そのまま)*

### 3.3. 予約確認（自分の予約一覧）

- **トリガー語:** 「確認」 / 「予約確認」 / 「予約状況」  
- **動作:** 予約ログから、現在のユーザー（名前・内線・職番のいずれか一致）に紐づく予約を最大 10 件まで一覧表示  
- **表示形式:**  
  `- <予約ID> [<ステータス>] <デモ機名> <開始日>→<終了日>`  
  例: `- 19fba056 [予約中] FE-01 2025-09-10→2025-09-12`
- **備考:**  
  - 一致条件は OR（名前一致、内線一致、職番一致のいずれか）  
  - 「予約中」以外のステータス（キャンセル済など）も表示対象  
  - 一覧は読み取り専用。キャンセルは「キャンセル」コマンドで実行  

## 4. フロントエンド設計
- **UI**: `index.html` + `static/main.js` によるシンプルなチャット画面。  
- **セッション管理**:  
  - ページ読み込み時に `sessionStorage` をリセットし常に新規会話を開始。  
  - `state`, `user_info`, `context` を保持し、各入力ごとに `POST /api/chat` に送信。  
  - API 応答で前記 3 つを更新し、レンダリング後スクロールを自動調整。  
- **初回メッセージ**: ページロード直後に自動で送信し、Bot の「お名前を教えてください。」から開始。  
- **エラー表示**: fetch エラー時には画面に System メッセージを追加。  

## 5. 開発環境
- **言語**: Python 3.13 （venv 同梱）  
- **主要ライブラリ**: Flask, openpyxl, python-dotenv, pytest  
- **Windows (COM モード時)**: 追加で pywin32 が必要。Excel がインストールされていること。  
- **テスト**: `python -m pytest -q` → 22 passed, 2 xfailed, 1 xpassed（Windows 固有回避あり）  
- **起動**:  
  ```powershell
  # 開発サーバー
  python src/app.py
  # COM モードで起動
  $env:EXCEL_WRITE_MODE = \"com\"
  python src/app.py
  ```  
- **運用注意**: Excel ファイルを Excel で開いたままにしない。サーバーは開放待ちするが、タイムアウトで明示エラーを返す。  
