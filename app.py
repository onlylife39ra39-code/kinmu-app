# app.py
import streamlit as st
import pandas as pd
import json
import requests
import re

from io_utils import export_excel
from engine import solve_schedule


# ============================
# HuggingFace AI 呼び出し
# ============================
HF_API_URL = "https://api-inference.huggingface.co/models/meta-llama/Llama-3-8B-Instruct"
HF_API_KEY = "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"  # ← まーくんのキーを入れる


def call_ai(prompt: str) -> str:
    headers = {"Authorization": f"Bearer {HF_API_KEY}"}
    payload = {
        "inputs": prompt,
        "parameters": {
            "max_new_tokens": 1024,
            "temperature": 0.2
        }
    }
    response = requests.post(HF_API_URL, headers=headers, json=payload)
    response.raise_for_status()
    data = response.json()

    # HuggingFaceの返答形式に合わせて抽出
    if isinstance(data, list) and len(data) > 0 and "generated_text" in data[0]:
        return data[0]["generated_text"]
    return str(data)


# ============================
# JSON抽出（安定化）
# ============================
def extract_json(text: str):
    """
    AIの返答から JSON 部分だけを安全に抜き出す。
    文章が混ざっても JSON を抽出できる。
    """
    json_match = re.search(r'\{[\s\S]*\}', text)
    if not json_match:
        raise ValueError("AI返答にJSONが見つかりませんでした。")

    json_text = json_match.group(0)

    try:
        return json.loads(json_text)
    except json.JSONDecodeError:
        raise ValueError("抽出したJSONが壊れています。プロンプトを調整してください。")


# ============================
# Streamlit UI
# ============================
st.set_page_config(page_title="勤務表自動生成（AI解析版）", layout="wide")
st.title("📘 勤務表自動生成システム（HuggingFace無料AI連携）")
st.sidebar.header("Excelアップロード")

uploaded_file = st.sidebar.file_uploader("主任作成の勤務表Excelをアップロード", type=["xlsx"])

if not uploaded_file:
    st.info("主任が作った勤務表Excelをそのままアップロードしてください。")
    st.stop()


# ============================
# Excel → テキスト化
# ============================
df_raw = pd.read_excel(uploaded_file, header=None)
text_data = "\n".join(df_raw[0].astype(str).tolist())

st.write("### Excel 1列目のテキスト化プレビュー")
st.text_area("テキスト", text_data, height=200)


# ============================
# AI 解析プロンプト
# ============================
prompt = f"""
You are an assistant that analyzes a Japanese nursing home work schedule Excel column.

Below is the first column of an Excel sheet used as a monthly work schedule
for care staff in a special nursing home.

Your task:
- Read the data as human-readable text.
- Extract structured information.

Extract the following:

1. Staff list:
  - name: staff full name (remove decorative symbols like ☆)
  - role: job title if present (e.g., 介護長, 介護主任)
  - group: group name if present (e.g., Aグループ, Bグループ, Cグループ)

2. Work codes:
  - List of unique work codes such as 日1, 日2, 公, 入浴, 有, 会議, etc.

Important:
- Exclude pure numbers (like 1, 0) from names.
- Exclude generic labels like グループ, 新入職員, パート.
- Exclude schedule description lines.
- Focus only on staff and work codes.

Output strictly in JSON with this structure:

{{
  "staff": [
    {{"name": "千明恵美", "role": "介護長", "group": null}},
    {{"name": "浦野裕太", "role": "介護主任", "group": null}},
    {{"name": "茂木最恵", "role": null, "group": "Aグループ"}}
  ],
  "codes": ["日1", "公", "入浴"]
}}

Do not add any explanation, only valid JSON.

Data:
{text_data}
"""

st.write("### AIに渡すプロンプト（確認用）")
with st.expander("プロンプトを見る"):
    st.code(prompt)


# ============================
# AI 解析実行
# ============================
if st.button("AIでExcelを解析する"):
    with st.spinner("AIがExcelを解析中..."):
        ai_response = call_ai(prompt)

        try:
            parsed = extract_json(ai_response)
        except Exception as e:
            st.error(f"AI返答の解析に失敗しました: {e}")
            st.text(ai_response)
            st.stop()

        staff_list = parsed.get("staff", [])
        codes_list = parsed.get("codes", [])

        if not staff_list:
            st.error("AIから職員情報が取得できませんでした。プロンプトの調整が必要です。")
            st.stop()

       
