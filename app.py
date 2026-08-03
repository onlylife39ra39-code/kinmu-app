import streamlit as st
import pandas as pd
import json
import re
import requests
import os
from io import BytesIO
from copy import deepcopy
from dotenv import load_dotenv

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Side, PatternFill, Font

# ============================================================
#  APIキー読み込み
# ============================================================
load_dotenv()
API_KEY = os.getenv("GEMINI_API_KEY")

if not API_KEY:
    st.error("❌ GEMINI_API_KEY が読み込めていません。.env を確認してください。")
    st.stop()

# ============================================================
#  Gemini REST API（3.6 Flash）
# ============================================================
def gemini_generate(prompt: str) -> str:
    url = (
        "https://generativelanguage.googleapis.com/v1beta/"
        "models/gemini-3.6-flash:generateContent?key=" + API_KEY
    )
    headers = {"Content-Type": "application/json"}
    data = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0, "maxOutputTokens": 8192},
    }
    response = requests.post(url, headers=headers, json=data)
    result = response.json()
    if "error" in result:
        st.error("Gemini API エラー:")
        st.json(result)
        st.stop()
    return result["candidates"][0]["content"]["parts"][0]["text"]

# ============================================================
#  勤務記号（文化OS ver3.0）
# ============================================================
CODE_LIST = [
    "公", "明公",
    "早1", "早A", "早B", "早C",
    "日1", "日2ホ",
    "遅1A", "遅1B", "遅1C",
    "夜1", "夜2"
]
# ============================================================
#  勤務記号入りシート自動判定
# ============================================================
def find_sheet_with_codes(dfs: dict) -> str | None:
    pattern = "|".join(CODE_LIST)
    for name, df in dfs.items():
        try:
            if df.apply(lambda row: row.astype(str).str.contains(pattern).any(), axis=1).any():
                return name
        except Exception:
            continue
    return None

# ============================================================
#  職員名判定（漢字・ひらがな・カタカナのみ）
# ============================================================
def is_staff_name(text: str) -> bool:
    if not text:
        return False
    t = text.replace("☆", "").strip()

    # 除外ワード
    if t in ["月間予定"]:
        return False
    if any(x in t for x in ["～", ":", "勤務時間", "週", "月～金", "土日祝"]):
        return False
    if any(x in t for x in ["グループ", "介護長", "介護主任", "新入職員", "パート"]):
        return False

    # 漢字・ひらがな・カタカナのみ
    if re.match(r'^[\u4E00-\u9FFF\u3040-\u309F\u30A0-\u30FF]+$', t):
        return True
    return False

# ============================================================
#  勤務記号抽出（確実に拾える安定版）
# ============================================================
def extract_codes(df_raw: pd.DataFrame):
    codes = set()
    for _, row in df_raw.iterrows():
        for cell in row:
            cell = str(cell).strip()
            if cell in CODE_LIST:
                codes.add(cell)
    return sorted(list(codes))
# ============================================================
#  役職・グループ抽出（まーくん勤務表構造完全対応版）
# ============================================================
def extract_roles_groups(df_raw: pd.DataFrame, filtered_names):
    roles = {}
    groups = {}

    current_role = None
    current_group = None

    for _, row in df_raw.iterrows():
        row_values = [str(x).strip() for x in row.tolist()]

        # 役職行
        if "介護長" in row_values:
            current_role = "介護長"
            current_group = None
            continue

        if "介護主任" in row_values:
            current_role = "介護主任"
            current_group = None
            continue

        if "新入職員" in row_values:
            current_role = "新入職員"
            current_group = None
            continue

        if "パート" in row_values:
            current_role = "パート"
            current_group = None
            continue

        # グループ行
        if "Aグループ" in row_values:
            current_group = "Aグループ"
            current_role = None
            continue

        if "Bグループ" in row_values:
            current_group = "Bグループ"
            current_role = None
            continue

        if "Cグループ" in row_values:
            current_group = "Cグループ"
            current_role = None
            continue

        # 職員名行
        for name in filtered_names:
            if name in row_values:
                if current_role:
                    roles[name] = current_role
                if current_group:
                    groups[name] = current_group

    return roles, groups

# ============================================================
#  既存勤務表の書式抽出（色・罫線・セル結合・列幅・行高さ）
# ============================================================
def extract_format_from_existing_excel(uploaded_file, sheet_name):
    # openpyxl で読み直す
    uploaded_file.seek(0)
    wb = load_workbook(uploaded_file)
    ws = wb[sheet_name]

    format_map = {}

    # セルごとの書式を全部吸い取る
    for row in ws.iter_rows():
        for cell in row:
            format_map[(cell.row, cell.column)] = {
                "fill": cell.fill,
                "border": cell.border,
                "font": cell.font,
                "alignment": cell.alignment,
            }

    # 列幅
    col_widths = {
        col: ws.column_dimensions[col].width
        for col in ws.column_dimensions
    }

    # 行高さ
    row_heights = {
        row: ws.row_dimensions[row].height
        for row in ws.row_dimensions
    }

    # セル結合
    merged_cells = list(ws.merged_cells.ranges)

    return format_map, col_widths, row_heights, merged_cells
# ============================================================
#  生成勤務表に既存書式を適用（完全コピー）
# ============================================================
def apply_format_to_generated_sheet(generated_data, format_map, col_widths, row_heights, merged_cells):
    wb = Workbook()
    ws = wb.active
    ws.title = "勤務表"

    # -------------------------
    # セル結合を完全コピー
    # -------------------------
    for mc in merged_cells:
        ws.merge_cells(str(mc))

    # -------------------------
    # 列幅を完全コピー
    # -------------------------
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    # -------------------------
    # 行高さを完全コピー
    # -------------------------
    for row, height in row_heights.items():
        ws.row_dimensions[row].height = height

    # -------------------------
    # 値＋書式を完全コピー
    # -------------------------
    for r_idx, row in enumerate(generated_data, start=1):
        for c_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)

            if (r_idx, c_idx) in format_map:
                fmt = format_map[(r_idx, c_idx)]

                # openpyxl の書式は deepcopy しないと TypeError
                cell.fill = deepcopy(fmt["fill"])
                cell.border = deepcopy(fmt["border"])
                cell.font = deepcopy(fmt["font"])
                cell.alignment = deepcopy(fmt["alignment"])

    return wb
# ============================================================
#  Streamlit UI（ページ設定）
# ============================================================
st.set_page_config(page_title="勤務表AI（完全統合版）", layout="wide")
st.title("📘 勤務表AI（Gemini 3.6 Flash / 本番仕様＋完全コピー）")

# ============================================================
#  Excel アップロード
# ============================================================
uploaded_file = st.sidebar.file_uploader("勤務表Excelをアップロード", type=["xlsx", "xlsm"])

if not uploaded_file:
    st.info("Excel をアップロードしてください。")
    st.stop()

# ============================================================
#  Excel 全シート読み込み
# ============================================================
xls = pd.ExcelFile(uploaded_file)
sheets = xls.sheet_names
dfs = {name: pd.read_excel(uploaded_file, sheet_name=name, header=None) for name in sheets}

st.write("### 読み込んだシート一覧")
st.json(sheets)

# ============================================================
#  勤務記号入りシートを自動判定
# ============================================================
target_sheet = find_sheet_with_codes(dfs)
if not target_sheet:
    st.error("❌ 勤務記号が含まれるシートが見つかりませんでした。")
    st.stop()

st.success(f"勤務記号入りシートを検出しました： {target_sheet}")

df_raw = dfs[target_sheet]

# ============================================================
#  職員名抽出
# ============================================================
name_col = df_raw.iloc[:, 1].fillna("").astype(str).tolist()
filtered_names = [n.replace("☆", "") for n in name_col if is_staff_name(n)]

st.write("### 抽出された職員名")
st.json(filtered_names)

# ============================================================
#  勤務記号抽出
# ============================================================
code_candidates = extract_codes(df_raw)

st.write("### 抽出された勤務記号（候補）")
st.json(code_candidates)

# ============================================================
#  役職・グループ抽出（まーくん勤務表構造）
# ============================================================
roles, groups = extract_roles_groups(df_raw, filtered_names)

st.write("### 役職候補")
st.json(roles)

st.write("### グループ候補")
st.json(groups)

# ============================================================
#  既存勤務表の書式抽出（色・罫線・セル結合・列幅・行高さ）
# ============================================================
format_map, col_widths, row_heights, merged_cells = extract_format_from_existing_excel(uploaded_file, target_sheet)
# ============================================================
# ① 既存勤務表の構造解析（staff / codes）
# ============================================================
st.write("## ① 既存勤務表の構造解析（staff / codes）")

if st.button("既存勤務表を解析する"):
    with st.spinner("Gemini が既存勤務表を解析中…"):

        analyze_prompt = f"""
あなたは勤務表解析AIです。
以下の Excel の構造に従って、職員の役職とグループを正確に割り当ててください。

【勤務表の構造】
- 役職とグループは「左端の列」に縦に並んでいる
- 上から順に以下の構造になっている：

介護長
介護主任
Aグループ（セル結合）
Bグループ（セル結合）
Cグループ（セル結合）
新入職員
パート（セル結合）

- 各役職・グループの直下に、その所属の職員名が縦に並ぶ
- 職員名は filtered_names に含まれる

【職員名一覧】
{json.dumps(filtered_names, ensure_ascii=False)}

【Excelから抽出した役職候補】
{json.dumps(roles, ensure_ascii=False)}

【Excelから抽出したグループ候補】
{json.dumps(groups, ensure_ascii=False)}

【勤務記号一覧（Excelから抽出）】
{json.dumps(code_candidates, ensure_ascii=False)}

出力形式（必ずこの形式で返す）：
{{
  "staff": [
    {{"name": "", "role": "", "group": ""}}
  ],
  "codes": []
}}

重要：
- JSON以外の文章は禁止
- 職員名は filtered_names のみ使用する
- 役職とグループは構造に従って必ず割り当てる
- 空欄を返さない
"""

        raw_output = gemini_generate(analyze_prompt)
        st.write("### Gemini 生出力（解析）")
        st.text(raw_output)

        try:
            json_text = re.search(r'\{[\s\S]*\}', raw_output).group(0)
            parsed_json = json.loads(json_text)
        except Exception as e:
            st.error(f"JSON解析に失敗しました: {e}")
            st.text(raw_output)
            st.stop()

        st.success("既存勤務表の解析が完了しました！")

        st.session_state["parsed_staff"] = parsed_json.get("staff", [])
        st.session_state["parsed_codes"] = parsed_json.get("codes", [])

        st.write("### staff（職員一覧）")
        st.json(st.session_state["parsed_staff"])

        st.write("### codes（勤務記号一覧）")
        st.json(st.session_state["parsed_codes"])
# ============================================================
# ② 翌月の勤務表を生成する（本番仕様JSON）
# ============================================================
st.write("## ② 翌月の勤務表を生成する（本番仕様JSON）")

default_month = "2026-06"
month = st.text_input("生成する月（YYYY-MM形式）", value=default_month)

if st.button("翌月の勤務表を生成する"):

    if "parsed_staff" not in st.session_state:
        st.warning("先に『既存勤務表を解析する』ボタンを押してください。")
        st.stop()

    with st.spinner("Gemini が勤務表を生成中…"):

        staff_json = st.session_state["parsed_staff"]
        codes_json = st.session_state["parsed_codes"] or code_candidates
        days = 30

        generate_prompt = f"""
JSONのみ返してください。

以下の仕様で勤務表を生成してください。

{{
  "month": "{month}",
  "days": {days},
  "staff": [
    {{
      "name": "",
      "role": "",
      "group": "",
      "schedule": {{
        "1": "",
        "2": "",
        "3": "",
        "4": "",
        "5": "",
        "6": "",
        "7": "",
        "8": "",
        "9": "",
        "10": "",
        "11": "",
        "12": "",
        "13": "",
        "14": "",
        "15": "",
        "16": "",
        "17": "",
        "18": "",
        "19": "",
        "20": "",
        "21": "",
        "22": "",
        "23": "",
        "24": "",
        "25": "",
        "26": "",
        "27": "",
        "28": "",
        "29": "",
        "30": ""
      }}
    }}
  ]
}}

【勤務記号一覧】
{json.dumps(codes_json, ensure_ascii=False)}

【職員一覧】
{json.dumps(staff_json, ensure_ascii=False)}

重要:
- JSON以外の文章は禁止
- 必ず完全な JSON を返す
- schedule は 1〜{days} まで全て埋める
"""

        raw_output = gemini_generate(generate_prompt)
        st.write("### Gemini 生出力（生成）")
        st.text(raw_output)

        try:
            json_text = re.search(r'\{[\s\S]*\}', raw_output).group(0)
            generated_schedule = json.loads(json_text)
        except Exception as e:
            st.error(f"JSON解析に失敗しました: {e}")
            st.text(raw_output)
            st.stop()

        st.success("翌月の勤務表が生成されました！")
        st.json(generated_schedule)

        st.session_state["generated_schedule"] = generated_schedule
# ============================================================
# ③ 生成した勤務表を既存勤務表の完全コピーでExcelに書き出す
# ============================================================
st.write("## ③ 生成した勤務表を既存勤務表の完全コピーでExcelに書き出す")

if "generated_schedule" in st.session_state:
    generated_schedule = st.session_state["generated_schedule"]
    staff_list = generated_schedule.get("staff", [])
    days = generated_schedule.get("days", 30)

    if staff_list:
        rows = []

        # 1行目ヘッダー
        header = ["職員名", "役職", "グループ"] + [str(d) for d in range(1, days + 1)]
        rows.append(header)

        # 2行目以降：職員ごとの行
        for staff in staff_list:
            row = [
                staff.get("name", ""),
                staff.get("role", ""),
                staff.get("group", "")
            ] + [
                staff.get("schedule", {}).get(str(d), "")
                for d in range(1, days + 1)
            ]
            rows.append(row)

        # Streamlit 表示用
        df_out = pd.DataFrame(rows[1:], columns=rows[0])
        st.write("### 生成勤務表（テーブル表示）")
        st.dataframe(df_out)

        # 既存勤務表の書式を完全コピーして新しいブックに適用
        wb = apply_format_to_generated_sheet(
            rows,
            format_map,
            col_widths,
            row_heights,
            merged_cells
        )

        buffer = BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        st.download_button(
            label="既存勤務表の完全コピーをExcelでダウンロード",
            data=buffer,
            file_name=f"勤務表_{generated_schedule.get('month', 'unknown')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("まだ勤務表が生成されていません。『既存勤務表を解析する』『翌月の勤務表を生成する』を先に実行してください。")
