# 職務経歴書 (Data Engineer / Backend Engineer)

## プロフェッショナルサマリー
Pythonを中心とした技術を用い、Web上の多様なソースからビジネス価値のあるデータを抽出・加工・統合する**データエンジニアリング**を専門としています。
単なるデータ収集にとどまらず、**ETL（Extract, Transform, Load）プロセスの自動化**、APIリクエスト最適化による**コスト削減**、およびGitHub ActionsやCloud Functionsを用いた**Serverlessな運用基盤の構築**を得意としています。
直近では、GeminiやClaude等のLLMをデータクレンジングや名寄せ処理に組み込み、非構造化データの品質を飛躍的に高める取り組みを行っています。

---

## 技術スタック (Technical Skills)

| カテゴリ | 技術・ツール |
| :--- | :--- |
| **Data Engineering** | **ETL Design**, Data Cleansing (Pandas/NumPy), Data Normalization, Format Conversion (JSON/CSV/Parquet) |
| **Languages** | **Python 3** (Main), Google Apps Script (GAS), SPARQL, VB.NET |
| **Cloud & Infra** | **Google Cloud** (Compute Engine, BigQuery, Cloud Functions), Linux (Ubuntu), Windows Server |
| **Automation & DevOps** | **GitHub Actions** (CI/CD, IssueOps), Docker, Task Scheduler, Cron |
| **API & Integration** | REST API Design, Google Maps API, YouTube Data API, BrightData, Slack/Chatwork API |
| **AI & LLM Ops** | **Prompt Engineering** (Gemini 1.5 Flash, Claude, ChatGPT), AI-based Data Correction |

---

## 主なプロジェクト実績 (Projects)

### 1. Google Maps API活用におけるデータ抽出最適化・分析基盤構築
**期間:** 2025年9月 ～ 2026年2月
**役割:** データエンジニア / バックエンドエンジニア
**環境:** Python 3, Google Cloud, GitHub Actions, Linux

**【プロジェクト概要】**
地理データを用いたマーケティング分析基盤の構築および運用コストの最適化。

**【データエンジニアリングとしての成果】**
* **APIコスト最適化 (Cost Optimization):**
    * 重複リクエストを排除するキャッシュシステムと、必要なフィールドのみを取得するクエリ設計により、API従量課金を大幅に削減。
* **運用自動化 (IssueOps):**
    * GitHub Actionsを活用し、Issueへのコメントをトリガーとして「データ抽出→加工→レポート生成」が走る自動化フローを構築。非エンジニアでも安全にバッチ処理を実行可能な環境を提供。
* **データ可視化:**
    * 取得した位置情報をヒートマップとして可視化し、エリア分析レポートを自動生成するパイプラインを実装。

---

### 2. Google Maps API データパイプライン プロトタイピング
**期間:** 2025年9月
**役割:** バックエンドエンジニア
**環境:** Python 3, Google Cloud, Gemini 1.5 Flash

**【業務内容】**
* 大規模データ抽出に耐えうるアーキテクチャの設計検証（PoC）。
* Gemini 1.5 Flashを用いた、抽出データの意味解析およびタグ付け精度の検証。

---

### 3. 学術研究用ナレッジグラフ構築 (Wikidata/MediaWiki)
**期間:** 2025年7月
**役割:** データエンジニア
**環境:** Python 3, SPARQL, Gemini 1.5 Flash

**【プロジェクト概要】**
学術利用を目的とした大規模公開データの整備および品質担保。

**【成果】**
* **大規模データ処理:** Wikidata APIおよびSPARQLクエリを駆使し、**11万件**のエンティティデータを効率的に抽出・統合。
* **AI品質管理 (AI-Driven QA):**
    * データの欠損や論理的矛盾を検知するため、Gemini 1.5 Flashを用いた自動チェック機構を実装。人手による確認コストを最小化した。

---

### 4. CRMデータベース更新バッチ開発 (Salesforce連携)
**期間:** 2025年6月 ～ 2025年7月
**役割:** バックエンドエンジニア
**環境:** Python 3, Google Sheets API, Excel

**【業務内容】**
* 企業データベースの鮮度維持を目的とした、データ更新バッチの開発。
* 複数のデータソースを照合し、Salesforce取り込み用のフォーマットへ自動変換するETL処理を実装。

---

### 5. 広告検証用プロキシネットワーク基盤構築
**期間:** 2025年3月 ～ 2025年4月
**役割:** インフラ/バックエンドエンジニア
**環境:** Python 3, BrightData API, Ubuntu

**【業務内容】**
* 広告配信の地域整合性を検証するため、BrightData APIを用いたIPローテーションおよびアクセスログ収集システムを構築。
* 取得したJSONログを解析可能な形式に変換し、データベースへ格納するフローを整備。

---

### 6. 地理データ統合・ID名寄せ
