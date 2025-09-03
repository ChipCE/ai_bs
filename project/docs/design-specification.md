# デモ機予約管理Bot 設計書

**設計書バージョン:** 2.1
**バージョン名:** 正確なExcel構造対応 最終版
**更新日時:** 2025.09.03

## 1. システムアーキテクチャ

### 1.1. コンポーネント構成
本システムは下図の4つのコンポーネントで構成され、Microsoft 365ライセンスの範囲内で動作します。

```mermaid
graph TD
    subgraph "1. UI層 (Microsoft Teams)"
        User[ユーザー]
    end

    subgraph "2. 対話エンジン層 (Power Platform)"
        PVA[Bot: Power Virtual Agents for Teams]
    end
    
    subgraph "3. ビジネスロジック層 (Power Platform)"
        subgraph "Power Automate Cloud Flows"
            Flow1[フロー1: 在庫確認]
            Flow2[フロー2: 予約確定]
            Flow3[フロー3: 予約リスト取得]
            Flow4[フロー4: 予約キャンセル]
        end
    end

    subgraph "4. データ層 (SharePoint Online)"
        Excel[Excel予約表]
    end

    User -- 対話 --> PVA;
    PVA -- 要求に応じて各フローを呼び出し --> Flow1;
    PVA -- 要求に応じて各フローを呼び出し --> Flow2;
    PVA -- 要求に応じて各フローを呼び出し --> Flow3;
    PVA -- 要求に応じて各フローを呼び出し --> Flow4;
    
    Flow1 -- スクリプト実行/セル読み取り --> Excel;
    Flow2 -- スクリプト実行/セル書き込み --> Excel;
    Flow3 -- テーブル読み取り --> Excel;
    Flow4 -- スクリプト実行/セル書き込み & テーブル行削除 --> Excel;

    Excel -- 処理結果 --> Flow1;
    Excel -- 処理結果 --> Flow2;
    Excel -- 処理結果 --> Flow3;
    Excel -- 処理結果 --> Flow4;
    
    Flow1 -- 結果を返す --> PVA;
    Flow2 -- 結果を返す --> PVA;
    Flow3 -- 結果を返す --> PVA;
    Flow4 -- 結果を返す --> PVA;

    PVA -- 応答メッセージ --> User;