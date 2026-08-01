import streamlit as st
import pandas as pd
import json
import re
import google.generativeai as genai
import os
from dotenv import load_dotenv

# -----------------------------
# Gemini APIキー設定（dotenv）
# -----------------------------
load_dotenv()  # .env を読み込む
genai.configure(api_key=os.getenv("GEMINI_API_KEY"))

model = genai.GenerativeModel("gemini-1.5-flash")

st.set_page_config(page_title="勤務表AI（Gemini自動化版）", layout="wide")
st.title("📘 勤務表AI（Gemini API 自動化版）")
st.write("Excel → JSON抽出 → 勤務表生成まで完全自動化します。")

# -----------------------------
# Excel アップロード
# -----------------------------
st.sidebar.header("Excelアップロード")
uploaded_file = st.sidebar.file_uploader("勤務表Excelをアップロード", type=["xlsx"])

if not uploaded_file:
    st.info("勤務表Excelをアップロードしてください。")
    st.stop()

df_raw = pd.read_excel(uploaded_file, header=None)

name_col = df_raw.iloc[:, 1].fillna("").astype(str).tolist()

# -----------------------------
# 職員名フィルタ
# -----------------------------
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
text_data = "\n".join(filtered_names)

st.write("### 抽出された職員名")
st.json(filtered_names)

# -----------------------------
# Gemini に送るプロンプト
# -----------------------------
prompt = f"""
あなたは介護施設の勤務表Excelを解析するAIです。

以下は主任が作成した勤務表Excelの「職員名が縦に並んだ列」です。
このデータから次を抽出してください：

1. staff（職員一覧）
  - name（職員名）
  - role（役職）
  - group（グループ）

2. codes（勤務記号一覧）

重要：
- 出力は JSON のみとし、説明文・Pythonコード・文章は一切含めないでください。
- JSONのトップレベルキーは必ず "staff" と "codes" の2つにしてください。

データ:
{text_data}
"""

# -----------------------------
# Gemini API 呼び出し（完全自動化）
# -----------------------------
st.write("### ▼ Gemini API による自動解析")
if st.button("勤務表を自動解析する"):
    with st.spinner("Gemini が勤務表を解析しています…"):
        response = model.generate_content(prompt)
        raw_output = response.text

        # JSON抽出
        try:
            m = re.search(r'\{[\s\S]*\}', raw_output)
            parsed_json = json.loads(m.group(0))
        except Exception as e:
            st.error(f"JSON解析に失敗しました: {e}")
            st.stop()

        st.success("JSON解析が完了しました！")

        st.write("### 職員一覧")
        st.json(parsed_json.get("staff", []))

        st.write("### 勤務記号一覧")
        st.json(parsed_json.get("codes", []))

        st.write("### ▼ ここから勤務表生成ロジックを追加できます（まーくんの既存コードを統合可能）")
