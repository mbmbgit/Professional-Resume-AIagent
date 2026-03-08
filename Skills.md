# 職務経歴書 (Data Engineer / Backend Engineer)

> 💡 **Professional Summary**
> **Python**を中心とした技術を用い、Web上の多様なソースからビジネス価値のあるデータを抽出・加工・統合する**データエンジニアリング**を専門としています。
> 単なるデータ収集にとどまらず、**ETLプロセスの自動化**、APIリクエスト最適化による**コスト削減**、およびGitHub ActionsやCloud Functionsを用いた**Serverlessな運用基盤の構築**を得意としています。
> 直近では、GeminiやClaude等の**LLM**をデータクレンジングや名寄せ処理に組み込み、非構造化データの品質を飛躍的に高める取り組みを行っています。

---

## 🛠 技術スタック (Technical Skills)

| Category | Tech Stack |
| :--- | :--- |
| **Data Engineering** | `ETL Design` `Data Cleansing (Pandas/NumPy)` `Data Normalization` `Format Conversion` |
| **Languages** | `Python 3` (Main) `Google Apps Script` `SPARQL` `VB.NET` |
| **Cloud & Infra** | **Google Cloud** (`Compute Engine` `BigQuery` `Cloud Functions`) `Linux (Ubuntu)` `Windows Server` |
| **DevOps & Auto** | `GitHub Actions` (CI/CD, IssueOps) `Docker` `Task Scheduler` `Cron` |
| **API Integration** | `REST API Design` `Google Maps API` `YouTube Data API` `BrightData` `Slack/Chatwork API` |
| **AI & LLM Ops** | `Gemini API` `Claude Code` `ChatGPT API` `Prompt Engineering` `AI-based Data Correction` |

---

## 📂 主なプロジェクト実績 (Projects)

### 🚀 1. Google Maps API活用におけるデータ抽出最適化・分析基盤構築
* **Period:** 2025/09 ～ 2026/02
* **Role:** データエンジニア / バックエンドエンジニア
* **Tech:** `Python 3` `Google Cloud` `GitHub Actions` `Linux` `Claude Code`

**【プロジェクト概要】**
地理データを用いたマーケティング分析基盤の構築および運用コストの最適化。

**【🏆 データエンジニアリングとしての成果】**
* **💰 APIコスト最適化 (Cost Optimization)**
    * 重複リクエストを排除するキャッシュシステムを実装。
    * 必要なフィールドのみを取得するクエリ設計により、**API従量課金を大幅に削減**。
* **🤖 運用自動化 (IssueOps)**
    * GitHub Actionsを活用し、Issueへのコメントをトリガーとして「データ抽出→加工→レポート生成」が走る自動化フローを構築。
    * 非エンジニアでも安全にバッチ処理を実行可能な環境を提供。
* **📊 データ可視化**
    * 取得した位置情報をヒートマップとして可視化し、エリア分析レポートを自動生成するパイプラインを実装。

---

### 🧪 2. Google Maps API データパイプライン プロトタイピング
* **Period:** 2025/09
* **Role:** バックエンドエンジニア
* **Tech:** `Python 3` `Google Cloud` `Gemini 1.5 Flash`

**【業務内容】**
* 大規模データ抽出に耐えうるアーキテクチャの設計検証（PoC）。
* `Gemini 1.5 Flash`を用いた、抽出データの意味解析およびタグ付け精度の検証。

---

### 📚 3. 学術研究用ナレッジグラフ構築 (Wikidata/MediaWiki)
* **Period:** 2025/07
* **Role:** データエンジニア
* **Tech:** `Python 3` `SPARQL` `Gemini 1.5 Flash`

**【プロジェクト概要】**
学術利用を目的とした大規模公開データの整備および品質担保。

**【成果】**
* **⚡ 大規模データ処理:** Wikidata APIおよびSPARQLクエリを駆使し、**11万件**のエンティティデータを効率的に抽出・統合。
* **✅ AI品質管理 (AI-Driven QA):** データの欠損や論理的矛盾を検知するため、Gemini 1.5 Flashを用いた自動チェック機構を実装。人手による確認コストを最小化した。

---

### 🔄 4. CRMデータベース更新バッチ開発 (Salesforce連携)
* **Period:** 2025/06 ～ 2025/07
* **Role:** バックエンドエンジニア
* **Tech:** `Python 3` `Google Sheets API` `Excel`

**【業務内容】**
* 企業データベースの鮮度維持を目的とした、データ更新バッチの開発。
* 複数のデータソースを照合し、Salesforce取り込み用のフォーマットへ自動変換するETL処理を実装。

---

### 🛡️ 5. 広告検証用プロキシネットワーク基盤構築
* **Period:** 2025/03 ～ 2025/04
* **Role:** インフラ/バックエンドエンジニア
* **Tech:** `Python 3` `BrightData API` `Ubuntu`

**【業務内容】**
* 広告配信の地域整合性を検証するため、BrightData APIを用いたIPローテーションおよびアクセスログ収集システムを構築。
* 取得したJSONログを解析可能な形式に変換し、データベースへ格納するフローを整備。

---

### 🗺️ 6. 地理データ統合・ID名寄せアプリケーション開発
* **Period:** 2024/11 ～ 2025/02
* **Role:** アプリケーションエンジニア
* **Tech:** `Python` `Flet` `Google Maps API` `Gemini 1.5 Flash`

**【データエンジニアリングとしての成果】**
* **🔗 データ名寄せ (Data Reconciliation):** 表記ゆれやプラットフォーム間で異なるID体系を持つデータを、LLM（Gemini）を用いて推論・統合するロジックを実装。
* **🖥️ GUI開発:** Python (Flet) を用いて、データ操作と結果確認を行えるデスクトップアプリを提供。

---

### 📝 7. 非構造化ドキュメント（論文）の構造化・要約システム
* **Period:** 2024/10
* **Role:** AIエンジニア
* **Tech:** `Python` `GCP` `Gemini API` `PyMuPDF`

**【業務内容】**
* PDFやHTML形式の学術論文からテキストデータを抽出（Extract）。
* Gemini APIを用いて要約およびメタデータ抽出を行い、検索可能なデータベースへ格納（Transform/Load）。

---

### 🔔 8. 競合サイトモニタリング・通知システム (Low-Code ETL)
* **Period:** 2023/05 ～ 2024/04
* **Role:** 開発・運用
* **Tech:** `Google Apps Script` `Octoparse` `Zapier`

**【業務内容】**
* 競合情報の更新検知から社内チャットへの通知までを、ノーコードツールとスクリプトを組み合わせて低コストに実現。

---

### 🏷️ 9. AI学習用データセット アノテーション
* **Period:** 2024/04 ～ 継続
* **Role:** データオペレーター
* **Tech:** `ELAN`

**【業務内容】**
* 動画データへのメタデータ付与（アノテーション）による教師データ作成。

---

### 🕷️ 10. Webデータ収集モジュール開発 (Python)
* **Period:** 2023/03 ～ 2024/02
* **Role:** Pythonエンジニア
* **Tech:** `Python 3` `Selenium`

**【業務内容】**
* Seleniumを用いた動的Webサイトからのデータ抽出モジュールの設計・実装。

---

### ⚙️ 11. データ収集ワークフロー設計 (Octoparse)
* **Period:** 2024/01 ～ 2024/02
* **Role:** ETLエンジニア
* **Tech:** `Octoparse` `CSV`

**【業務内容】**
* データ収集ツールを用いた定期実行ワークフローの設計およびデータ出力設定。

---

### 🤝 12. データ収集ツール導入コンサルティング
* **Period:** 2023/12 ～ 2024/01
* **Role:** テクニカルサポート
* **Tech:** `Octoparse`

**【業務内容】**
* クライアント企業へのツール導入支援および技術レクチャー。

---

### 🧹 13. データクレンジングおよび正規化処理
* **Period:** 2023/09 ～ 2023/10
* **Role:** データエンジニア
* **Tech:** `Python 3` `Pandas`

**【業務内容】**
* 収集した生データ（Raw Data）に対し、Pandasを用いて欠損値処理・型変換・正規化を行い、分析用データセットとして納品。

---

### 📄 14. OCR活用デスクトップアプリ開発 (DX支援)
* **Period:** 2023/05 ～ 2023/06
* **Role:** アプリケーションエンジニア
* **Tech:** `Python` `LINE CLOVA OCR` `ChatGPT API`

**【業務内容】**
* 紙帳票のデジタル化を支援するOCRアプリ開発。API連携およびUXを考慮したGUI設計を担当。

---

### 🤖 15. 業務自動化スクリプト開発 (Python)
* **Period:** 2023/05 ～ 2023/06
* **Role:** Pythonエンジニア
* **Tech:** `Python 3` `ChatGPT API`

**【業務内容】**
* 定型業務を自動化するPythonスクリプトの作成および展開。

---

### 📊 16. ダッシュボード用データパイプライン構築 (GCP)
* **Period:** 2023/03 ～ 2023/04
* **Role:** クラウドエンジニア
* **Tech:** `GCE (Ubuntu)` `Python` `GAS`

**【成果】**
* BIツールへのデータ供給を自動化するため、GCE上での定期データ収集・加工ジョブを構築。サーバーレス環境への移行を見据えた設計を実施。

---

### 🏛️ 17. 行政データ クレンジング・リスト生成 (17万件)
* **Period:** 2023/02 ～ 2023/03
* **Role:** データエンジニア
* **Tech:** `Python` `Regex`

**【業務内容】**
* 公開されている行政データ（約17万件）に対し、正規表現を用いた高度なクレンジング処理を実施。住所・電話番号の表記ゆれを統一し、CRMへインポート可能な形式に整形。

---

### 🖼️ 18. 画像収集・管理ツール開発
* **Period:** 2022/12 ～ 2023/01
* **Role:** Pythonエンジニア
* **Tech:** `Python 3` `GUI Lib`

**【業務内容】**
* マーケティング素材収集の効率化を目的とした画像収集・保存ツールの開発。

---

### 🌍 19. 海外ゲームデータ収集 RPA (UiPath)
* **Period:** 2022/12
* **Role:** RPAエンジニア
* **Tech:** `UiPath` `VB.NET`

**【業務内容】**
* 海外サイトからのデータ収集プロセスをRPAツールで自動化。

---

### 🏗️ 20. オブジェクト指向を用いたクローリングシステム設計
* **Period:** 2022/09 ～ 2023/10
* **Role:** Pythonエンジニア
* **Tech:** `Python 3`

**【成果】**
* 大規模かつ長期的なデータ収集に耐えうるよう、クラス設計（ポリモーフィズム）を採用した保守性の高いクローラーを構築。仕様変更への対応工数を削減。

---

### 🎥 21. YouTube Data API連携リスト生成
* **Period:** 2021/10 ～ 2021/12
* **Role:** Pythonエンジニア
* **Tech:** `Python` `YouTube Data API`

**【業務内容】**
* YouTube Data APIを活用し、特定条件に合致するインフルエンサーリストを自動生成するツールを開発。

---

### 👁️ 22. サイト更新検知システム (Selenium)
* **Period:** 2021/08 ～ 2021/09
* **Role:** Pythonエンジニア
* **Tech:** `Python` `Selenium` `Chatwork API`

**【業務内容】**
* Webサイトの更新を検知し、Chatwork APIを通じて即時通知を行う監視システムの実装。

---

### 🏢 23. データセンター運用保守 (Infrastructure Ops)
* **Period:** 2017/10 ～ 2018/04
* **Role:** インフラオペレーター
* **Tech:** `Windows Server` `Systemwalker` `UNIX`

**【業務内容】**
* サーバー・ネットワーク機器の監視、バックアップ運用、障害一次対応（エスカレーション）。

---
