"""
md_to_excel.py
--------------
Skills.md を読み込み、以下を一括実行する:
  1. SkillSheet_Manpyo.csv を上書き生成
  2. SkillSheet_Manpyo.xlsx を書式付きで生成

プロジェクト数は Skills.md に動的に対応。

実行方法:
    pip install openpyxl
    python3 md_to_excel.py
"""

import re
import csv
from csv_to_excel import build_excel  # Excel生成ロジックを流用

INPUT_MD    = "Skills.md"
OUTPUT_CSV  = "SkillSheet_Manpyo.csv"
OUTPUT_XLSX = "SkillSheet_Manbyo.xlsx"

# ── Skills.md に記載のないプロフィール項目（必要に応じて編集） ──
PROFILE_STATIC = {
    "所属":     "個人事業主（フリーランス）",
    "稼動開始": "4/13-",
    "資格":     "ITパスポート",
    "得意技術": "Python / Selenium / Octoparse / Gemini API / Claude Code / GCP / GitHub Actions / GitHub Pages / GitHub Issues / Flet / GAS / FastAPI / BigQuery / Firebase / Docker / AWS EC2",
    "得意業務": "業務自動化（RPA・スクレイピング・データ収集） / AI連携システム開発 / GUIアプリ開発 / Web監視・通知システム",
    "年齢":     "42",
    "性別":     "男",
    "最寄駅":   "JR甲子園口",
    "学歴":     "立命館大学法学部政治行政専攻 卒業",
}

COMPANIES = [
    {
        "no": "1", "name": "個人事業主（フリーランス）",
        "period": "2019年12月 - 現在", "type": "個人事業主",
        "role": "Pythonエンジニア", "business": "Webシステム開発・業務自動化・AI連携",
    },
]


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# テキスト整形ユーティリティ
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def clean_md(text):
    """Markdown装飾（**bold**, `code`）を除去してプレーンテキスト化。"""
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
    text = re.sub(r'`(.+?)`', r'\1', text)
    return text.strip()


def clean_detail_text(text):
    """業務内容・効果テキストを整形（Markdownリスト → 全角箇条書き）。"""
    lines = []
    for line in text.split('\n'):
        line = line.rstrip()
        line = re.sub(r'\*\*(.+?)\*\*', r'\1', line)   # 太字除去
        line = re.sub(r'^\*\s+', '　・', line)          # * bullet
        line = re.sub(r'^-\s+', '　・', line)           # - bullet
        line = re.sub(r'^    ', '　　', line)            # 4sp indent → 全角インデント
        lines.append(line)
    # 先頭・末尾の空行を除去
    while lines and not lines[0].strip():
        lines.pop(0)
    while lines and not lines[-1].strip():
        lines.pop()
    return '\n'.join(lines)


def extract_between(text, start_pattern, end_pattern):
    """start_pattern 〜 end_pattern 間のテキストを返す。end が見つからなければ末尾まで。"""
    m = re.search(start_pattern, text)
    if not m:
        return ''
    start_pos = m.end()
    m2 = re.search(end_pattern, text[start_pos:])
    if m2:
        return text[start_pos: start_pos + m2.start()]
    return text[start_pos:]


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 期間・役割パーサー
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def parse_period(period_str):
    """
    期間文字列 → (start, end, duration)
    例:
      "2024年11月 〜 2026年2月（約1年4ヶ月）" → ("2024年11月", "2026年2月", "約1年4ヶ月")
      "2025年7月（約1ヶ月）"                  → ("2025年7月", "2025年7月", "約1ヶ月")
      "2026年3月 〜 2026年4月"               → ("2026年3月", "2026年4月", "")
    """
    m = re.search(r'(\d{4}年\d+月)\s*[〜~]\s*(\d{4}年\d+月)(?:（(.+?)）)?', period_str)
    if m:
        return m.group(1), m.group(2), m.group(3) or ''
    m = re.search(r'(\d{4}年\d+月)(?:（(.+?)）)?', period_str)
    if m:
        return m.group(1), m.group(1), m.group(2) or ''
    return period_str, '', ''


def parse_role(role_str):
    """
    "PG・PM / 1名" → ("PG・PM", "1名")
    "アノテーター / 1名" → ("アノテーター", "1名")
    """
    m = re.match(r'(.+?)\s*/\s*(\d+名|\d+人)', role_str)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return role_str.strip(), '1名'


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Skills.md パーサー
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def parse_project_block(block, seq_no):
    """
    1プロジェクト分の Markdown ブロックをパースして辞書を返す。
    戻り値のキーは csv_to_excel.parse_csv の出力と同じ形式。
    """
    # タイトル
    m = re.match(r'### \d+\.\s*(.+)', block)
    if not m:
        return None
    title = '■' + m.group(1).strip()

    # --- 情報テーブル（| **Key** | Value |）をパース ---
    info = {}
    for line in block.split('\n'):
        tm = re.match(r'\|\s*\*\*(.+?)\*\*\s*\|\s*(.+?)\s*\|$', line)
        if tm:
            key = tm.group(1).strip()
            val = clean_md(tm.group(2).strip())
            if '期間' in key:
                info['期間'] = val
            elif '役割' in key:
                info['役割'] = val
            elif '使用言語' in key:
                info['使用言語'] = val
            elif key == 'DB':
                info['DB'] = val
            elif 'サーバOS' in key:
                info['サーバOS'] = val
            elif 'FW' in key:
                info['FW'] = val

    # --- 期間・役割パース ---
    start, end, duration = parse_period(info.get('期間', ''))
    role, size = parse_role(info.get('役割', ''))

    # --- 技術スタック ---
    lang   = info.get('使用言語', '-')
    db     = info.get('DB', '-')
    os_str = info.get('サーバOS', '-')
    fw_raw = info.get('FW', '-')
    # カンマ区切り → 改行区切りに変換
    fw = '\n'.join(x.strip() for x in fw_raw.split(','))

    # --- 業務内容・ビジネス効果テキスト ---
    detail_raw = extract_between(
        block,
        r'\*\*業務内容・詳細機能:\*\*',
        r'\*\*ビジネス上の効果:\*\*'
    )
    biz_raw = extract_between(
        block,
        r'\*\*ビジネス上の効果:\*\*',
        r'\| 要件定義'
    )

    parts = []
    if detail_raw.strip():
        parts.append('≪担当業務≫\n' + clean_detail_text(detail_raw))
    if biz_raw.strip():
        parts.append('≪ビジネス上の効果≫\n' + clean_detail_text(biz_raw))
    detail = '\n\n'.join(parts)

    # --- 担当工程テーブル ---
    phases = []
    for line in block.split('\n'):
        if '|' in line and ':---:' not in line:
            cells = re.findall(r'\|\s*(●|-)\s*', line)
            if len(cells) == 7:
                phases = cells
                break
    if not phases:
        phases = ['-'] * 7

    return {
        'no':       str(seq_no),
        'start':    start,
        'end':      end,
        'title':    title,
        'role':     f'{role}\n{size}',
        'lang':     lang,
        'db':       db,
        'os':       os_str,
        'fw':       fw,
        'phases':   phases,
        'detail':   detail,
        'team':     f'体制：{size}\n（個人事業主）',
        'duration': f'（{duration}）' if duration else '',
    }


def parse_md(filename):
    """
    Skills.md をパースして (profile, companies, projects) を返す。
    戻り値は csv_to_excel.parse_csv と同じ形式。
    """
    with open(filename, encoding='utf-8') as f:
        content = f.read()

    # プロフィール（氏名のみ MD から取得、残りは PROFILE_STATIC）
    profile = dict(PROFILE_STATIC)
    m = re.search(r'\*\*氏名\*\*:\s*(.+)', content)
    profile['氏名'] = m.group(1).strip() if m else '萬俵 秀俊（まんぴょう ひでとし）'

    # プロジェクト経歴セクションを抽出
    proj_start = content.find('## ■ プロジェクト経歴')
    if proj_start == -1:
        proj_start = content.find('## ■ プロジェクト')
    proj_end = re.search(r'\n## ■(?!.* プロジェクト)', content[proj_start + 1:])
    proj_content = content[proj_start: proj_start + 1 + proj_end.start()] \
        if proj_end else content[proj_start:]

    # ### N. Title ブロックに分割
    blocks = re.split(r'\n(?=### \d+\. )', proj_content)

    projects = []
    seq = 1
    for block in blocks:
        if not re.match(r'### \d+\. ', block):
            continue
        proj = parse_project_block(block, seq)
        if proj:
            projects.append(proj)
            seq += 1

    return profile, COMPANIES, projects


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# CSV ライター
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

NCOLS = 19  # A〜S列


def make_row(*pairs):
    """(列インデックス, 値) のペアで 19列リストを生成。"""
    row = [''] * NCOLS
    for col, val in pairs:
        row[col] = val
    return row


def write_csv(profile, companies, projects, filename):
    rows = []

    # ── プロフィールセクション ───────────────────────────────────
    rows.append(make_row((0, 'スキルシート')))
    rows.append(make_row((0, '技術者名'), (1, profile['氏名']),
                         (8, '所　　属'), (9, profile['所属'])))
    rows.append(make_row((0, '年　　齢'), (8, '性　　別')))
    rows.append(make_row((0, '最 寄 駅'), (8, '学　　歴')))
    rows.append(make_row((0, '稼動開始日'), (1, profile['稼動開始']),
                         (8, '資　　格'),  (9, profile['資格'])))
    rows.append(make_row((0, '得意技術'), (1, profile['得意技術'])))
    rows.append(make_row((0, '得意業務'), (1, profile['得意業務'])))

    # ── 職務経歴ヘッダー ─────────────────────────────────────────
    rows.append(make_row((0, '職務経歴'), (2, 'No.'), (3, '企業名'),
                         (6, '期間'), (7, '契約・雇用形態'),
                         (9, '担当職種'), (11, '事業内容')))
    for comp in companies:
        rows.append(make_row((2, comp['no']), (3, comp['name']),
                             (6, comp['period']), (7, comp['type']),
                             (9, comp['role']), (11, comp['business'])))

    # ── プロジェクトセクションヘッダー ───────────────────────────
    rows.append(make_row((0, '期間'), (5, '業務内容'),
                         (7, '役割\n規模'), (8, '使用言語'), (9, 'DB'),
                         (10, 'サーバOS'), (11, 'FW・MW\nツール等'),
                         (12, '担当工程')))
    rows.append(make_row((12, '要件定義'), (13, '基本設計'), (14, '詳細設計'),
                         (15, '実装・単体'), (16, '結合テスト'),
                         (17, '総合テスト'), (18, '保守・運用')))

    # ── プロジェクトデータ（3行/件） ─────────────────────────────
    for i, proj in enumerate(projects):
        # 行+0: タイトル行
        title_row = [''] * NCOLS
        title_row[0]  = str(i + 1)
        title_row[1]  = proj['start']
        title_row[3]  = '-'
        title_row[4]  = proj['end']
        title_row[5]  = proj['title']
        title_row[7]  = proj['role']
        title_row[8]  = proj['lang']
        title_row[9]  = proj['db']
        title_row[10] = proj['os']
        title_row[11] = proj['fw']
        for j, ph in enumerate(proj['phases']):
            title_row[12 + j] = ph
        rows.append(title_row)

        # 行+1: 詳細テキスト行
        rows.append(make_row((5, proj['detail']), (7, proj['team'])))

        # 行+2: 期間行
        rows.append(make_row((1, proj['duration'])))

    with open(filename, 'w', newline='', encoding='utf-8-sig') as f:
        csv.writer(f).writerows(rows)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# メイン
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def main():
    # Step 1: Skills.md をパース
    print(f"1. {INPUT_MD} を読み込み中...")
    profile, companies, projects = parse_md(INPUT_MD)
    print(f"   プロジェクト数: {len(projects)} 件")
    for p in projects:
        print(f"   - {p['no']:>2}. {p['title'][:40]}")

    # Step 2: CSV 出力
    print(f"\n2. {OUTPUT_CSV} を生成中...")
    write_csv(profile, companies, projects, OUTPUT_CSV)
    print(f"   ✓ {OUTPUT_CSV}")

    # Step 3: Excel 出力（csv_to_excel.build_excel を流用）
    print(f"\n3. {OUTPUT_XLSX} を生成中...")
    wb = build_excel(profile, companies, projects)
    wb.save(OUTPUT_XLSX)
    print(f"   ✓ {OUTPUT_XLSX}")

    print("\n✅ 完了")


if __name__ == "__main__":
    main()
