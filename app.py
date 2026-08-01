import streamlit as st
import pandas as pd
import json
import re
import requests
import os
from dotenv import load_dotenv

load_dotenv()
API_KEY = os.getenv("GEMINI_API_KEY")

if not API_KEY:
    st.error("❌ GEMINI_API_KEY が読み込めていません。")
    st.stop()

def gemini_generate(prompt):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key={API_KEY}"

    headers = {"Content-Type": "application/json"}
    data = {
        "contents": [
            {"parts": [{"text": prompt}]}
        ]
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

st.set_page_config(page_title="勤務表AI（Gemini 3.6 Flash）", layout="wide")
st.title("📘 勤務表AI（Gemini 3.6 Flash REST API）")

uploaded_file = st.sidebar.file_uploader("勤務表Excelをアップロード", type=["xlsx"])

if not uploaded_file:
    st.info("Excel をアップロードしてください。")
    st.stop()

df_raw = pd.read_excel(uploaded_file, header=None)
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
text_data = "\n".join(filtered_names)

st.write("### 抽出された職員名")
st.json(filtered_names)

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
