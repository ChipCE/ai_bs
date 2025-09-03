
---

### 2. 設計書 (`docs/design-specification.md`)

```markdown
# デモ機予約管理Bot 設計書

**設計書バージョン:** 5.0
**バージョン名:** Chatbotステートマシン実装版
**更新日時:** 2025.09.03

## 1. システムアーキテクチャ

### 1.1. 概要
ユーザーのPC上で完結するWebアプリケーションを構築する。Chatbotの対話ロジックは、バックエンドのPythonサーバーに実装されたステートマシンによって管理される。フロントエンドは、現在の対話状態（State）を保持し、ユーザーの入力をサーバーに送信する役割を担う。

### 1.2. コンポーネント構成

```mermaid
graph TD
    subgraph "Local PC"
        subgraph "User Interface (Web Browser)"
            Frontend[HTML/CSS/JavaScript]
        end

        subgraph "Application Server"
            PythonServer[Python Web Server (Flask/FastAPI)]
        end

        subgraph "Data Storage"
            Excel[Local Excel File]
        end

        Frontend -- API Call (HTTP) --> PythonServer
        PythonServer -- Read/Write --> Excel
    end