"""
generate_proposal.py
--------------------
案件募集文（.txt）と Skills.md をもとに Gemini API で提案文を生成し、
proposals/proposal_<yyyymmdd_HHMMSS>.md として保存する。

使い方:
    python3 generate_proposal.py <案件ファイル.txt> [--policy "方針テキスト"] [--template-file <テンプレートファイル.md>]

    例:
        python3 generate_proposal.py job_postings/fooma.txt
        python3 generate_proposal.py job_postings/fooma.txt --policy "800件PDF案件の実績を特に強調すること"
        python3 generate_proposal.py job_postings/fooma.txt --template-file templates/short_template.md

    または .txt ファイル末尾に以下のセクションを記載しても同様に動作する:

        ===方針===
        ・800件PDF案件の実績を特に強調すること
        ・文章を短めに収めること

環境変数:
    GEMINI_API_KEY  : Gemini API キー（必須）

GitHub Actions では secrets.GEMINI_API_KEY を環境変数として渡す。
"""

import argparse
import json
import os
import re
import sys
import urllib.request
import urllib.error
from datetime import datetime
from pathlib import Path

# ─────────────────────────────────────────
# 定数
# ─────────────────────────────────────────
SKILLS_MD         = Path("Skills.md")
TEMPLATE_MD       = Path("proposal_template.md")
OUTPUT_DIR        = Path("proposals")
GEMINI_ENDPOINT   = (
    "https://generativelanguage.googleapis.com/v1beta/models/"
    "gemini-2.5-pro:generateContent"
)
MIN_HOURLY_RATE   = 1500  # 円/時：これを下回ると警告を表示する


# ─────────────────────────────────────────
# ファイル読み込み
# ─────────────────────────────────────────
def load_file(path: Path) -> str:
    if not path.exists():
        sys.exit(f"[ERROR] ファイルが見つかりません: {path}")
    return path.read_text(encoding="utf-8")


# ─────────────────────────────────────────
# 案件テキストから方針・テンプレート指示セクションを分離する
#   ===方針=== 以降を方針として取り出し、
#   ===テンプレート=== 以降をテンプレート指示として取り出す
# ─────────────────────────────────────────
def split_policy(text: str) -> tuple[str, str, str]:
    """(案件本文, 方針, テンプレート指示) のタプルを返す。各セクションがなければ空文字。"""
    # ===テンプレート=== セクションを分離
    tmpl_sep = re.split(r"^===テンプレート===$", text, maxsplit=1, flags=re.MULTILINE)
    if len(tmpl_sep) == 2:
        text, template_hint = tmpl_sep[0].strip(), tmpl_sep[1].strip()
    else:
        template_hint = ""
    # ===方針=== セクションを分離
    sep = re.split(r"^===方針===$", text, maxsplit=1, flags=re.MULTILINE)
    if len(sep) == 2:
        return sep[0].strip(), sep[1].strip(), template_hint
    return text.strip(), "", template_hint


# ─────────────────────────────────────────
# Gemini API 呼び出し（リトライ付き）
# ─────────────────────────────────────────
def call_gemini(api_key: str, prompt: str, max_retries: int = 3) -> str:
    url = f"{GEMINI_ENDPOINT}?key={api_key}"
    payload = json.dumps({
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.7,
            "maxOutputTokens": 4096,
            "topP": 0.95,
        },
    }).encode("utf-8")
    headers = {"Content-Type": "application/json"}

    for attempt in range(1, max_retries + 1):
        print(f"[Gemini] API呼び出し {attempt}/{max_retries} ...")
        req = urllib.request.Request(url, data=payload, headers=headers, method="POST")
        try:
            with urllib.request.urlopen(req, timeout=60) as res:
                body = json.loads(res.read().decode("utf-8"))
                text = body["candidates"][0]["content"]["parts"][0]["text"]
                if text:
                    return text
        except urllib.error.HTTPError as e:
            if attempt == max_retries:
                sys.exit(f"[ERROR] Gemini API エラー: {e.code} {e.reason}")
            print(f"  → HTTPError {e.code}。10秒後にリトライ...")
            import time; time.sleep(10)
        except (KeyError, IndexError) as e:
            sys.exit(f"[ERROR] レスポンス解析エラー: {e}")

    sys.exit("[ERROR] Gemini API から有効な回答を取得できませんでした。")


# ─────────────────────────────────────────
# 単価・作業時間を推定して時給をチェックする
# ─────────────────────────────────────────
def analyze_hourly_rate(job_posting: str, api_key: str) -> dict:
    """案件文から単価と作業時間を推定し、結果を辞書で返す。失敗時は空辞書。"""
    prompt = f"""以下の案件募集文を分析してください。

【案件募集文】
{job_posting}

次の項目を推定し、JSONのみを出力してください（説明文・コードブロック記法は不要）。
金額・時間が不明または記載なしの場合は null を入れてください。

{{
  "price_min": <最低報酬額（円・整数・不明ならnull）>,
  "price_max": <最高報酬額（円・整数・不明ならnull）>,
  "hours_min": <最低作業時間（時間・小数可・不明ならnull）>,
  "hours_max": <最高作業時間（時間・小数可・不明ならnull）>,
  "price_note": "<単価の補足（単発/月次/要相談など）>",
  "hours_note": "<作業時間の補足>"
}}"""

    url = f"{GEMINI_ENDPOINT}?key={api_key}"
    payload = json.dumps({
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.1, "maxOutputTokens": 512},
    }).encode("utf-8")
    headers = {"Content-Type": "application/json"}

    req = urllib.request.Request(url, data=payload, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=30) as res:
            body = json.loads(res.read().decode("utf-8"))
            text = body["candidates"][0]["content"]["parts"][0]["text"]
            m = re.search(r"\{.*\}", text, flags=re.DOTALL)
            if m:
                return json.loads(m.group())
    except Exception as e:
        print(f"[警告] 時給分析APIの呼び出しに失敗しました: {e}")
    return {}


def print_hourly_rate_check(analysis: dict, auto_yes: bool = False) -> None:
    """推定単価・作業時間・時給を表示し、基準を下回る場合は警告を出す。"""
    price_min  = analysis.get("price_min")
    price_max  = analysis.get("price_max")
    hours_min  = analysis.get("hours_min")
    hours_max  = analysis.get("hours_max")
    price_note = analysis.get("price_note", "")
    hours_note = analysis.get("hours_note", "")

    print("─" * 52)
    print("【単価・時給チェック】")

    # 単価表示
    if price_min is not None or price_max is not None:
        if price_min == price_max or price_max is None:
            print(f"  単価       : {price_min:,}円  ※{price_note}")
        elif price_min is None:
            print(f"  単価       : 〜{price_max:,}円  ※{price_note}")
        else:
            print(f"  単価       : {price_min:,}〜{price_max:,}円  ※{price_note}")
    else:
        print(f"  単価       : 不明（{price_note}）")

    # 作業時間表示
    if hours_min is not None or hours_max is not None:
        if hours_min == hours_max or hours_max is None:
            print(f"  作業時間   : 約{hours_min}時間  ※{hours_note}")
        elif hours_min is None:
            print(f"  作業時間   : 〜{hours_max}時間  ※{hours_note}")
        else:
            print(f"  作業時間   : {hours_min}〜{hours_max}時間  ※{hours_note}")
    else:
        print(f"  作業時間   : 不明（{hours_note}）")

    # 時給計算（最悪ケース：最低単価÷最大時間）
    warned = False
    if price_min is not None and hours_max is not None and hours_max > 0:
        hourly_worst = price_min / hours_max
        print(f"  推定時給   : 約{hourly_worst:,.0f}円/時（最悪ケース: 最低単価÷最大時間）")
        if hourly_worst < MIN_HOURLY_RATE:
            print(f"\n  [警告] 推定時給 {hourly_worst:,.0f}円/時 が基準 {MIN_HOURLY_RATE:,}円/時 を下回っています！")
            warned = True

    # 最良ケース（最大単価÷最小時間）
    if price_max is not None and hours_min is not None and hours_min > 0:
        hourly_best = price_max / hours_min
        print(f"  推定時給   : 約{hourly_best:,.0f}円/時（最良ケース: 最高単価÷最小時間）")

    if warned:
        # CI 環境 (GitHub Actions 等) または --yes フラグ指定時は自動続行
        is_ci = bool(os.environ.get("CI"))
        if auto_yes or is_ci:
            tag = "CI自動続行" if is_ci else "--yes"
            print(f"  → [{tag}] 警告を確認しましたが生成を続行します。")
        else:
            answer = input("  → 提案文の生成を続けますか？ [y/N]: ").strip().lower()
            if answer not in ("y", "yes"):
                sys.exit("[中断] 時給が基準を下回るため提案文の生成を中止しました。")

    print("─" * 52 + "\n")


# ─────────────────────────────────────────
# プロンプト構築
# ─────────────────────────────────────────
def build_prompt(job_posting: str, skills_md: str, template: str, policy: str = "", template_hint: str = "") -> str:
    policy_section = ""
    if policy:
        policy_section = f"""
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【提案文の方針・追加指示】
以下の指示を最優先で守って提案文を生成してください。
{policy}
"""

    is_free_mode = not template.strip()

    if is_free_mode:
        # テンプレートが空 → Gemini が構成を自由に作成
        task_line = "案件内容とスキルシートをもとに、最終的な提案文をMarkdown形式で自由に作成してください。"
        rules = """【出力ルール】
- 伝え方の順序は1.Point（結論）: 私の強みは〇〇です。2.Reason（理由）: なぜなら、〇〇という経験をしたからです。3.Example（具体例）: 具体的には、FastAPIを使って〜（ここで技術の話）。4.Point（結論）: だから、貴社で貢献できます。
- 提案文は日本語で、クライアントに送付できる丁寧なビジネス文体にしてください。
- 読みやすく自然な構成で、最適なセクション・見出しを自由に設定してください。
- 箇条書き形式で■や*などを使用して、読みやすくしてください。
- 「メタ情報」セクションは出力に含めないでください。
- 全体として読みやすく、押しつけがましくない文体にしてください。
- 各文章は簡潔に説明してください。"""
        template_section = """━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【テンプレート】
テンプレートの指定がありません。案件内容をもとに最適な構成で自由に作成してください。
参考構成例: 挨拶・応募の一文 → 関連実績（2〜3件） → 本案件への対応方針（3〜5文） → プロフィール・締め"""
    else:
        # テンプレートあり → {{PLACEHOLDER}} を埋める従来モード
        task_line = "以下の【案件募集文】と【スキルシート】をよく読み、【テンプレート】の各 {{{{PLACEHOLDER}}}} を埋めた最終的な提案文をMarkdown形式で出力してください。"
        rules = """【出力ルール】
- 提案文は日本語で、クライアントに送付できる丁寧なビジネス文体にしてください。
- テンプレートの構成（セクション順・見出し）は維持してください。
- 「メタ情報」セクションは出力に含めないでください。
- {{{{CASE_TITLE}}}} には案件の簡潔なタイトルを入れてください。
- {{{{DATE}}}} には今日の日付（yyyy年m月d日）を入れてください。
- {{{{OPENING_LINE}}}} には案件名を踏まえた応募の一文（1〜2文）を入れてください。
- {{{{RELEVANT_EXPERIENCE}}}} には【スキルシート】の中から本案件に関連する実績を2〜3件選び、  
  各実績を1〜2文で簡潔に説明してください。箇条書きリスト形式（- ）を使ってください。
- {{{{APPROACH}}}} には案件の要件・注意点を踏まえた具体的な対応方針を3〜5文で記述してください。  
  表・箇条書きは使わず、自然な文章にしてください。
- 全体として読みやすく、押しつけがましくない文体にしてください。
- 本案件への対応方針は次の流れで書いてください。Point（結論）: 私の強みは〇〇です。⇒Reason（理由）: なぜなら、〇〇という経験をしたからです。⇒Example（具体例）: 具体的には、FastAPIを使って〜（ここで技術の話）。⇒Point（結論）: だから、貴社で貢献できます。
"""
        template_section = f"""━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【テンプレート】
{template}"""

    # テンプレート追加指示（--template または ===テンプレート=== で指定）
    hint_section = ""
    if template_hint:
        hint_section = f"""
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【テンプレートへの追加指示】（最優先で反映してください）
{template_hint}
"""

    return f"""あなたは優秀なフリーランスエンジニアの提案文ライターです。
{task_line}

{rules}
{policy_section}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【案件募集文】
{job_posting}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【スキルシート】
{skills_md}

{template_section}
{hint_section}"""


# ─────────────────────────────────────────
# メタ情報ヘッダーを除去して本文だけ返す
# ─────────────────────────────────────────
def strip_meta(text: str) -> str:
    return re.sub(r"<!-- META:.*?<!-- /META -->", "", text, flags=re.DOTALL).strip()


# ─────────────────────────────────────────
# エントリーポイント
# ─────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description="案件募集文から提案文を Gemini API で自動生成する"
    )
    parser.add_argument("job_posting_file", help="案件ファイルのパス（.txt）")
    parser.add_argument(
        "--policy",
        default="",
        help="提案文の生成方針を文字列で指定（例: '800件PDF案件を特に強調すること'）",
    )
    parser.add_argument(
        "--template",
        default="",
        help="テンプレートへの追加指示を文字列で指定（テンプレートファイルが空の場合も有効）",
    )
    parser.add_argument(
        "--template-file",
        default="",
        metavar="FILE",
        help=f"使用するテンプレートファイルのパスを指定（省略時は {TEMPLATE_MD} を使用）",
    )
    parser.add_argument(
        "--yes",
        action="store_true",
        help="時給が低くても確認なしで続行する（GitHub Actions など非対話環境用）",
    )
    args = parser.parse_args()

    api_key = os.environ.get("GEMINI_API_KEY", "")
    if not api_key:
        sys.exit("[ERROR] 環境変数 GEMINI_API_KEY が設定されていません。")

    job_posting_path = Path(args.job_posting_file)
    raw_posting      = load_file(job_posting_path)
    skills_md        = load_file(SKILLS_MD)
    template_path    = Path(args.template_file) if args.template_file else TEMPLATE_MD
    template         = load_file(template_path)
    if args.template_file:
        print(f"[テンプレートファイル] {template_path}")

    # ファイル内の ===方針=== / ===テンプレート=== セクションを分離
    job_posting, file_policy, file_template_hint = split_policy(raw_posting)

    # --policy オプションを優先し、なければファイル内方針を使う
    policy = args.policy or file_policy
    if policy:
        print(f"[方針] {policy[:80]}{'...' if len(policy) > 80 else ''}")

    # --template オプションを優先し、なければファイル内テンプレート指示を使う
    template_hint = args.template or file_template_hint
    if template_hint:
        print(f"[テンプレート指示] {template_hint[:80]}{'...' if len(template_hint) > 80 else ''}")

    # テンプレートが空の場合は自由生成モードを通知
    if not template.strip():
        print("[テンプレート] テンプレートが空のため、Gemini が自由に構成を作成します。")

    # 単価・時給チェック（提案文生成前に実施）
    print("[分析] 単価・作業時間を推定して時給をチェックしています...")
    analysis = analyze_hourly_rate(job_posting, api_key)
    if analysis:
        print_hourly_rate_check(analysis, auto_yes=args.yes)
    else:
        print("[警告] 単価・時給の分析結果が取得できませんでした。スキップします。\n")

    prompt   = build_prompt(job_posting, skills_md, template, policy, template_hint)
    raw_text = call_gemini(api_key, prompt)

    # 出力ファイル名
    OUTPUT_DIR.mkdir(exist_ok=True)
    timestamp   = datetime.now().strftime("%Y%m%d_%H%M%S")
    stem        = job_posting_path.stem
    output_path = OUTPUT_DIR / f"proposal_{stem}_{timestamp}.md"

    output_path.write_text(raw_text, encoding="utf-8")
    print(f"[OK] 提案文を保存しました: {output_path}")


if __name__ == "__main__":
    main()
