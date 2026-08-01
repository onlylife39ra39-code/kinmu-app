import streamlit as st
import pandas as pd
import json
import re
import requests
import os
from dotenv import load_dotenv

# -----------------------------
# APIキー読み込み
# -----------------------------
load_dotenv()
API_KEY = os.getenv("GEMINI_API_KEY")

if not API_KEY:
    st.error("❌ GEMINI_API_KEY が読み込めていません。")
    st.stop()

# -----------------------------
# Gemini REST API（3.6 Flash 安定版）
# -----------------------------
def gemini_generate(prompt):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key={API_KEY}"

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

    try:
        result = response.json()
    except Exception:
        st.error("❌ API が JSON 以外の応答を返しました")
        st.write(response.text)
        st.stop()

    if "error" in result:
        st.error("Gemini API エラー:")
        st.json(result)
        st.stop()

    return result["candidates"][0]["content"]["parts"][0]["text"]

# -----------------------------
# 勤務記号入りシートを自動判定
# -----------------------------
def find_sheet_with_codes(dfs):
    code_keywords = ["早", "遅", "夜", "休", "公休", "有給", "日勤", "研修", "出張"]

    for name, df in dfs.items():
        try:
            if df.apply(lambda row: row.astype(str).str.contains("|".join(code_keywords)).any(), axis=1).any():
                return name
        except:
            continue
    return None

# -----------------------------
# Streamlit UI
# -----------------------------
st.set_page_config(page_title="勤務表AI（Gemini 3.6 Flash / シート自動判定版）", layout="wide")
st.title("📘 勤務表AI（Gemini 3.6 Flash / シート自動判定版）")

uploaded_file = st.sidebar.file_uploader("勤務表Excelをアップロード", type=["xlsx"])

if not uploaded_file:
    st.info("Excel をアップロードしてください。")
    st.stop()

# 全シート読み込み
xls = pd.ExcelFile(uploaded_file)
sheets = xls.sheet_names
dfs = {name: pd.read_excel(uploaded_file, sheet_name=name, header=None) for name in sheets}

# 勤務記号入りシートを探す
target_sheet = find_sheet_with_codes(dfs)

if not target_sheet:
    st.error("❌ 勤務記号が含まれるシートが見つかりませんでした。")
    st.write("読み込んだシート一覧:", sheets)
    st.stop()

st.success(f"勤務記号入りシートを検出しました： {target_sheet}")

df_raw = dfs[target_sheet]

# -----------------------------
# 職員名抽出
# -----------------------------
name_col = df_raw.iloc[:, 1].fillna("").astype(str).tolist()

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

filtered_names = [n.replace("☆", "") for n in name_col if is_staff_name(n)]
st.write("### 抽出された職員名")
st.json(filtered_names)

# -----------------------------
# 勤務記号抽出（行全体から）
# -----------------------------
all_rows = df_raw.fillna("").astype(str).values.tolist()

code_candidates = set()
for row in all_rows:
    for cell in row:
        if re.match(r"^(早|遅|夜|休|公休|有給|日勤|研修|出張)$", cell):
            code_candidates.add(cell)

st.write("### 抽出された勤務記号（候補）")
st.json(list(code_candidates))

# -----------------------------
# Gemini プロンプト（最適化）
# -----------------------------
prompt = f"""
あなたは介護施設の勤務表Excelを解析するAIです。

以下の職員名と勤務記号候補から、次の JSON を作成してください：

1. staff（職員一覧）
  - name（職員名）
  - role（役職）
  - group（グループ）

2. codes（勤務記号一覧）

重要：
- 出力は JSON のみ。
- JSON以外の文章は一切含めない。
- 必ず完全な JSON を返す（途中で切らない）。
- JSONはできるだけ短く簡潔にする。

職員名:
{filtered_names}

勤務記号候補:
{list(code_candidates)}
"""

# -----------------------------
# REST API による自動解析
# -----------------------------
st.write("### ▼ Gemini 3.6 Flash による自動解析")

if st.button("勤務表を自動解析する"):
    with st.spinner("Gemini が解析中…"):
        raw_output = gemini_generate(prompt)

        try:
            m = re.search(r'\{[\s\S]*\}', raw_output)
            parsed_json = json.loads(m.group(0))
        except Exception as e:
            st.error(f"JSON解析に失敗しました: {e}")
            st.write("Gemini出力:", raw_output)
            st.stop()

        st.success("JSON解析が完了しました！")

        st.write("### 職員一覧")
        st.json(parsed_json.get("staff", []))

        st.write("### 勤務記号一覧")
        st.json(parsed_json.get("codes", []))
