import streamlit as st
import pandas as pd
import json
import re
import requests
import os
from dotenv import load_dotenv

# ============================================================
#  APIキー読み込み
# ============================================================
load_dotenv()
API_KEY = os.getenv("GEMINI_API_KEY")

if not API_KEY:
    st.error("❌ GEMINI_API_KEY が読み込めていません。.env を確認してください。")
    st.stop()

# ============================================================
#  Gemini REST API（3.6 Flash 安定版）
# ============================================================
def gemini_generate(prompt: str) -> str:
    url = (
        "https://generativelanguage.googleapis.com/v1beta/"
        "models/gemini-3.6-flash:generateContent?key=" + API_KEY
    )

    headers = {"Content-Type": "application/json"}

    data = {
        "contents": [
            {"parts": [{"text": prompt}]}
        ],
        "generationConfig": {
            "temperature": 0,
            "maxOutputTokens": 8192
        }
    }

    response = requests.post(url, headers=headers, json=data)
    result = response.json()

    if "error" in result:
        st.error("Gemini API エラー:")
        st.json(result)
        st.stop()

    return result["candidates"][0]["content"]["parts"][0]["text"]

# ============================================================
#  勤務記号（まーくんの実データ）
# ============================================================
CODE_LIST = [
    "公", "明公",
    "早1", "早A", "早B", "早C",
    "日1", "日2ホ",
    "遅1A", "遅1B", "遅1C",
    "夜1", "夜2"
]

# ============================================================
#  勤務記号入りシートを自動判定
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
#  職員名判定
# ============================================================
def is_staff_name(text: str) -> bool:
    if not text:
        return False
    t = text.replace("☆", "").strip()

    if t in ["月間予定"]:
        return False
    if any(x in t for x in ["～", ":", "勤務時間", "週", "月～金", "土日祝"]):
        return False
    if any(x in t for x in ["グループ", "介護長", "介護主任", "新入職員", "パート"]):
        return False

    if re.match(r'^[\u4E00-\u9FFF\u3040-\u309F\u30A0-\u30FF]+$', t):
        return True
    return False

# ============================================================
#  Streamlit UI
# ============================================================
st.set_page_config(page_title="勤務表AI（統合版）", layout="wide")
st.title("📘 勤務表AI（Gemini 3.6 Flash / 統合版）")

uploaded_file = st.sidebar.file_uploader("勤務表Excelをアップロード", type=["xlsx", "xlsm"])

if not uploaded_file:
    st.info("Excel をアップロードしてください。")
    st.stop()

# ============================================================
#  全シート読み込み
# ============================================================
xls = pd.ExcelFile(uploaded_file)
sheets = xls.sheet_names
dfs = {name: pd.read_excel(uploaded_file, sheet_name=name, header=None) for name in sheets}

st.write("### 読み込んだシート一覧")
st.json(sheets)

# ============================================================
#  勤務記号入りシートを探す
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
all_rows = df_raw.fillna("").astype(str).values.tolist()

code_candidates = set()
for row in all_rows:
    for cell in row:
        if cell in CODE_LIST:
            code_candidates.add(cell)

st.write("### 抽出された勤務記号（候補）")
st.json(sorted(list(code_candidates)))

# ============================================================
#  ① 既存勤務表の解析（staff / codes）
# ============================================================
st.write("## ① 既存勤務表の構造解析（staff / codes）")

if st.button("既存勤務表を解析する"):
    with st.spinner("Gemini が既存勤務表を解析中…"):

        analyze_prompt = f"""
JSONのみ返してください。

{{ 
 "staff": [
   {{"name": "", "role": "", "group": ""}}
 ],
 "codes": {json.dumps(sorted(list(code_candidates)), ensure_ascii=False)}
}}
"""

        raw_output = gemini_generate(analyze_prompt)

        st.write("### Gemini 生出力（デバッグ用）")
        st.text(raw_output)

        try:
            m = re.search(r'\{[\s\S]*\}', raw_output)
            parsed_json = json.loads(m.group(0))
        except Exception as e:
            st.error(f"JSON解析に失敗しました: {e}")
            st.write("Gemini出力:", raw_output)
            st.stop()

        st.success("既存勤務表の解析が完了しました！")

        st.session_state["parsed_staff"] = parsed_json.get("staff", [])
        st.session_state["parsed_codes"] = parsed_json.get("codes", [])

        st.write("### staff（職員一覧）")
        st.json(st.session_state["parsed_staff"])

        st.write("### codes（勤務記号一覧）")
        st.json(st.session_state["parsed_codes"])

# ============================================================
#  ② 翌月勤務表生成（文化OS ver3.0 ロジック統合）
# ============================================================
st.write("## ② 翌月の勤務表を生成する（文化OSロジック統合）")

default_month = "2026-06"
month = st.text_input("生成する月（YYYY-MM形式）", value=default_month)

if st.button("翌月の勤務表を生成する"):

    if "parsed_staff" not in st.session_state:
        st.warning("先に『既存勤務表を解析する』ボタンを押してください。")
        st.stop()

    with st.spinner("Gemini が勤務表を生成中…"):

        staff_json = st.session_state["parsed_staff"]
        codes_json = st.session_state["parsed_codes"]

        generate_prompt = f"""
JSONのみ返してください。

{{
 "month": "{month}",
 "staff": [
   {{
     "name": "",
     "role": "",
     "group": "",
     "schedule": {{
       "1": "",
       "2": "",
       "3": ""
     }}
   }}
 ]
}}
"""

        raw_output = gemini_generate(generate_prompt)

        st.write("### Gemini 生出力（デバッグ用）")
        st.text(raw_output)

        try:
            m = re.search(r'\{[\s\S]*\}', raw_output)
            generated_schedule = json.loads(m.group(0))
        except Exception as e:
            st.error(f"JSON解析に失敗しました: {e}")
            st.write("Gemini出力:", raw_output)
            st.stop()

        st.success("翌月の勤務表が生成されました！")
        st.json(generated_schedule)
