# job_postings/

案件の募集文テキストファイルをここに置いてください。

## 使い方

1. 案件の募集文を `.txt` ファイルとしてこのフォルダに保存します。
   - ファイル名例: `fooma_2026.txt`, `scraping_202604.txt`
2. `main` ブランチに push すると、GitHub Actions が自動で提案文を生成します。
3. 生成された提案文は `proposals/` フォルダに保存されます。

## ローカルで実行する場合

```bash
GEMINI_API_KEY=your_api_key python3 generate_proposal.py job_postings/fooma_2026.txt
```
