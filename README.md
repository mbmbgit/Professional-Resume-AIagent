## 🤖 AI Issue Responder (Demo)

このリポジトリでは、GitHub ActionsとGemini APIを連携させた **AI自動応答システム** を稼働させています。
*This repository features an **AI Auto-Response System** powered by GitHub Actions and the Gemini API.*

### 🛠 使い方と仕様 (How it Works)

1. **Issueを作成またはコメントする (Create or Comment on an Issue)**
   - Issueを立てるか、既存のIssueにコメントをすると、私のAIアシスタント（萬俵AI）が自動で返信します。
   - *My AI assistant will automatically respond when you open a new issue or post a comment.*
   - [https://github.com/mbmbgit/Professional-Resume-AIagent/issues/1](https://github.com/mbmbgit/Professional-Resume-AIagent/issues/1)

2. **多言語対応 (Multilingual Support)**
   - 日本語で入力すると日本語で、英語で入力すると英語で回答します。ユーザーが使用した言語をAIが自動で判別し、最適な言語で回答します。
   - *Inputs in Japanese will be answered in Japanese, and English inputs will be answered in English. The AI automatically detects your language and responds accordingly.*

3. **応答時間 (Response Time)**
   - GitHub Actionsの起動とAIの処理を含め、返信まで **通常15秒ほど** かかります。
   - *It usually takes about **15 seconds** for the AI to respond, including the workflow trigger and processing time.*

### 💡 開発の背景 (Purpose)

このシステムは、単なる自動応答ボットではありません。プロンプトの微調整（Prompt Engineering）による回答精度の向上や、安全性（Safety Settings）の確保、そしてGitHub Actionsを用いた効率的なIssue-opsの実践を目的として研究・運用しています。

*This system is not just an automated bot. It is a playground for **Prompt Engineering** to improve response accuracy, ensure safety settings, and practice efficient **Issue-ops** using GitHub Actions.*

---

# 職務経歴書

**現在日**: 2026年3月23日現在  
**氏名**: 萬俵 秀俊 (まんぴょう ひでとし)

---

## ■ 職務要約
**「自動化による業務効率化」と「顧客満足度の高い納品」に強みを持つエンジニア**

2019年よりフリーランスのエンジニアとして、Pythonを用いたWeb開発・スクレイピング・業務自動化案件に従事。クラウドソーシング等のプラットフォームを通じ、中小企業を中心に180件以上のプロジェクトを完遂。顧客満足度は97%を維持しています。
要件定義から開発、納品後のサポートまで一貫して対応できる「完結力」と、リサーチ時間を90%削減するなどの「課題解決力」を活かし、即戦力として貢献いたします。

---

## ■ 活かせるスキル・知識

| 分野 | 詳細 |
| :--- | :--- |
| **言語・FW** | Python (FastAPI), JavaScript (React), Go, C++, HTML/CSS, VB.NET, GAS |
| **インフラ・DB** | GCP, Docker, Linux (Ubuntu), Windows Server, Selenium, GitHub Actions, IssueOps |
| **AI・ツール** | 生成AI活用 (Gemini, Claude Code), Octoparse, UiPath, Flet |
| **資格・知識** | ITパスポート |

---

## ■ 主要な職務経歴（IT・エンジニアリング）

### 個人事業主（フリーランス エンジニア）
* **期間**: 2019年12月 ～ 現在
* **事業内容**: Webシステム開発、RPA導入支援、リサーチ業務

**【実績要約】**
* **総受注数**: 174件 （総プロジェクト数 183件）
* **顧客評価**: 満足度97% （高評価 178件 / 総評価 183件）

**【主な取り組みと成果】**
* **業務効率化ツールの開発（Python / Selenium）**
  * **課題**: 顧客のリサーチ業務が手作業で膨大な時間を要していた。
  * **対策**: スクレイピングと自動集計ツールをPythonで開発し、ワークフローを構築。
  * **成果**: 作業時間を**10時間 → 1時間（90%削減）**へ短縮。
* **クラウドソーシングでの高評価獲得**
  * **工夫**: 曖昧な要件に対しても、プロトタイプを早期に提示する「提案型」のコミュニケーションを徹底。
  * **成果**: ランサーズ等のプラットフォームにてリピート受注を含む多数の高評価を獲得し、平均報酬単価の向上を実現。

---

## ■ プロジェクト経歴（詳細スキルシート）
個人事業主（クラウドソーシング等）として担当した主な開発実績です。すべては書ききれませんでした。
（※特記がない限り、すべて1名体制 / プログラマとしての参画です）

| 期間 | プロジェクト名 / 業務内容 | 担当フェーズ | 技術環境 (言語 / OS / DB / ツール等) |
| :--- | :--- | :--- | :--- |
| 2025/09～2026/02 | **GoogleMap API リクエスト数削減のためのデータ分析フロー追加**<br>効率的なデータ抽出システム開発、ヒートマップ作成、分析レポート提出およびシステム運用 | 要件定義〜テスト | Python 3, Linux, CSV<br>GCP, GitHub Actions, Codespaces, Claude Code |
| 2025/09～2025/09 | **GoogleMap APIを活用したデータ抽出フロー構築**<br>APIを用いた効率的な抽出システムの開発 | 要件定義〜テスト | Python 3, Linux, CSV<br>GCP, GitHub Actions, Gemini 1.5 Flash |
| 2025/07～2025/07 | **Wikipediaデータの抽出**<br>学術用データ整備。APIを利用し11万件の抽出・クレンジング、AIによる論理的品質チェック | 要件定義〜テスト | Python 3, SPARQL, Windows, CSV<br>VS Code, Gemini 1.5 Flash |
| 2025/06～2025/07 | **Salesforce用の企業データベース更新**<br>既存の企業データの更新および自動化 | 要件定義〜テスト | Python 3, Windows, Excel<br>Google Spreadsheet, Gemini 1.5 Flash |
| 2025/03～2025/04 | **プロキシサービス導入支援**<br>広告検証のためのPythonでのAPI連携プログラム開発 | 詳細設計・テスト | Python, Ubuntu, JSON<br>BrightData API, GCP console, Windsurf |
| 2024/11～2025/02 | **GoogleMap APIデータ抽出ソフトウェア開発**<br>データ抽出およびfletでのGUI開発。表記ゆれや別プラットフォームのIDをAIで統合 | 要件定義〜テスト | Python, Windows, CSV<br>Flet, Place Detail API, Gemini 1.5 Flash, Windsurf |
| 2024/10～2024/10 | **論文要約のためのAIシステム開発**<br>PDF/HTMLのテキスト変換、テキスト要約AIの自動化プログラム開発 | 基本設計〜テスト | Python, Linux, Google Spreadsheet<br>GCP, Gemini API, pyMuPDF, BeautifulSoup |
| 2024/04～2024/05 | **AIアノテーション**<br>動画にアノテーションを付与する作業 | 担当 | ELAN |
| 2024/01～2024/02 | **OCRシステム開発**<br>データ収集ソフトウェアのワークフロー開発 | 詳細設計〜テスト | Octoparse, CSV |
| 2023/12～2024/01 | **Webスクレイピングツールの使用レクチャー**<br>データ収集ソフトウェアの開発オンラインレクチャー | 導入・サポート | Octoparse, CSV |
| 2023/05～2024/04 | **Web監視通知システム開発 (ノーコード)**<br>Web監視通知システムの開発・実装 | 詳細設計〜テスト | GAS, Windows, Google Spreadsheet<br>Octoparse, Zapier, Chatwork |
| 2023/03～2024/02 | **Webスクレイピングコード作成**<br>データ収集プログラムの開発 | 詳細設計〜テスト | Python 3, VS Code<br>Selenium 4.1 |
| 2023/09～2023/10 | **Webスクレイピングコード作成**<br>データ収集、データクレンジング | 詳細設計〜テスト | Python 3, VS Code |
| 2023/05～2023/06 | **OCRシステム開発**<br>APIを使用したOCRデスクトップアプリ開発。GUI開発、API連携、ローカルデータ連携、UX提案 | 基本設計〜導入 | Python 3, Windows 10, CSV<br>Line Clova OCR, VS Code, ChatGPT |
| 2023/05～2023/06 | **RPAワークフロー開発**<br>データ収集プログラム開発、データクレンジング | 基本設計〜導入 | Python 3, Windows 10, Spreadsheet<br>VS Code, ChatGPT |
| 2023/03～2023/04 | **RPAワークフロー開発**<br>ダッシュボード作成用データ収集プログラム。GASでのプログラム実装 | 基本設計〜導入 | Python 3, GAS, Windows 10/Ubuntu, CSV<br>タスクスケジューラ, Chatwork, VS Code, ChatGPT |
| 2023/02～2023/03 | **RPAワークフロー開発**<br>DM送付用リスト作成。行政データ(17万件)のリサーチ・クレンジング、独自ドメインURL抽出 | 要件定義〜導入 | Python 3, Ubuntu, CSV<br>メールソフト, 筆まめ, VS Code |
| 2022/12～2023/01 | **RPAネイティブアプリ開発**<br>マーケティング用データ収集。スクレイピング、スクショ保存自動化、GUI実装 | 基本設計〜導入 | Python 3, Windows 10, Excel<br>VS Code |
| 2022/12～2022/12 | **RPAワークフロー開発**<br>海外ゲームデータ収集（クローリングプログラム開発） | 基本設計〜導入 | VB.NET, Windows 10, Excel<br>UiPath |
| 2022/09～2023/10 | **RPAワークフロー開発**<br>分析用データ収集。GUI実装、classを使ったポリモーフィズム設計 | 要件定義〜導入 | Python 3, Windows 10, CSV<br>VS Code |
| 2021/10～2021/12 | **RPAワークフロー開発**<br>リスト作成プログラム開発 | 要件定義〜導入 | Python, Windows 10, CSV<br>YouTube Data API, VS Code |
| 2021/08～2021/09 | **RPAワークフロー開発**<br>Web監視システム。Seleniumを利用した監視プログラム | 要件定義〜導入 | Python, Windows 10, CSV<br>Chatwork, VS Code |
| 2019/08～2019/12 | **AI画像・広告アノテーション**<br>建築現場の画像にアノテーションを付与 | 作業担当 | Windows 10 |
| 2018/12～2019/01 | **RPAワークフロー開発**<br>スクレイピングワークフローの作成・導入 | 要件定義〜導入 | Windows 10, CSV<br>Octoparse |

---

## ■ その他の職務経歴
ITエンジニアとしての業務と並行し、製造・物流業界にて現場の生産性向上や品質管理業務に従事してまいりました。

| 期間 | 雇用形態 | 派遣先・業務内容 |
| :--- | :--- | :--- |

| 2023年7月～2023年12月 | 業務委託 | **民泊運営代行会社**<br>客室清掃および設備点検。詳細な報告連絡により期間中無事故を達成。 |
| 2018年10月～2018年12月 | 派遣社員 | **株式会社Adecco**（FUJITSU系列）<br>データセンターオペレーション（サーバー監視、媒体交換、環境構築支援など）。体調不良により、退職 |
| 2017年2月～2018年12月 | 業務委託 | **西日本マリンサービス株式会社**<br>派遣スタッフとしてAmazon西宮FCでピッキングスタッフ。生産性向上に具体的な施設設置案を提出⇒採用 |
| 2007年4月～2018年9月 | 派遣・正社員 | **製造・物流各社に派遣スタッフとして勤務**（トヨタ自動車九州、ROHM,松下、黒井電機、など）<br>生産ラインでの組立、梱包、品質チェック業務。<br>※厳しい品質基準が求められる現場にて目標生産数を達成。 |

---

## ■ 自己PR

**1. フェイルセーフのエンジニアリング**
単にコードを書くだけでなく、「いかに効率よく動かすか」に注力しています。前述のリサーチ業務自動化（10時間→1時間短縮）の事例のように、PythonやGitHub等の技術を組み合わせ、コスト削減と効率化という具体的なビジネス成果を提供できます。

**2. リモートワークでの自律的なプロジェクト遂行力**
5年以上のフリーランス経験を通じて、対面ではない環境下でも信頼関係を築くコミュニケーション能力を磨きました。テキストベースでの正確な要件確認、進捗の可視化、納期前の納品を徹底しており、どのような環境でも自律的に業務を遂行できます。

**3. 新技術への適応と継続的な学習**
未経験から独学と実践でスキルを習得してきました。現在は生成AI（LLM）を活用し、市場価値の高い技術を積極的に取り入れています。GitHub等でのアウトプットも継続しており、変化の激しいIT業界においても柔軟に適応し続けます。

---

## ■ ポートフォリオ・リンク

* **GitHub**: [https://github.com/mememori8888/Professional-Resume-AIagent/issues/1](https://github.com/mememori8888/Professional-Resume-AIagent/issues/1)
* **SkillSheet**: [Google Spreadsheets](https://docs.google.com/spreadsheets/d/1r68Yx4GzO0NEgAKrwx_ocUjKWZ7wbbHEJUMpMEkelWY/edit?usp=sharing)
* **Lancers Profile**: [https://www.lancers.jp/profile/hideman_mememori8888](https://www.lancers.jp/profile/hideman_mememori8888)
