import streamlit as st
import pandas as pd
import json
import re
import requests
import os
from io import BytesIO
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
st.set_page_config(page_title="勤務表AI（本番仕様統合版）", layout="wide")
st.title("📘 勤務表AI（Gemini 3.6 Flash / 本番仕様＋Excel出力）")

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

        st.write("### Gemini 生出力（解析）")
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
#  ② 翌月勤務表生成（本番仕様JSON）
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
        codes_json = st.session_state["parsed_codes"]

        # 月の日数はとりあえず30で固定（あとで動的にしてもOK）
        days = 30

        generate_prompt = f"""
JSONのみ返してください。

以下の仕様で勤務表を生成してください。

【JSON仕様】
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

重要：
- JSON以外の文章は禁止。
- 必ず完全な JSON を返す。
- schedule は 1〜{days} まで全て埋める。
"""

        raw_output = gemini_generate(generate_prompt)

        st.write("### Gemini 生出力（生成）")
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

        st.session_state["generated_schedule"] = generated_schedule

# ============================================================
#  ③ 生成した勤務表をExcelに書き出す
# ============================================================
st.write("## ③ 生成した勤務表をExcelに書き出す")

if "generated_schedule" in st.session_state:
    generated_schedule = st.session_state["generated_schedule"]
    staff_list = generated_schedule.get("staff", [])
    days = generated_schedule.get("days", 30)

    # DataFrame化：行＝職員、列＝日付
    if staff_list:
        rows = []
        for staff in staff_list:
            row = {"職員名": staff.get("name", ""), "役職": staff.get("role", ""), "グループ": staff.get("group", "")}
            schedule = staff.get("schedule", {})
            for d in range(1, days + 1):
                row[str(d)] = schedule.get(str(d), "")
            rows.append(row)

        df_out = pd.DataFrame(rows)

        st.write("### 生成勤務表（テーブル表示）")
        st.dataframe(df_out)

        # Excel書き出し
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            df_out.to_excel(writer, index=False, sheet_name="勤務表")

        buffer.seek(0)

        st.download_button(
            label="生成した勤務表をExcelでダウンロード",
            data=buffer,
            file_name=f"勤務表_{generated_schedule.get('month', 'unknown')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("まだ勤務表が生成されていません。『翌月の勤務表を生成する』を先に実行してください。")
