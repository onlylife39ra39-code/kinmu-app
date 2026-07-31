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

    if isinstance(data, list) and len(data) > 0 and "generated_text" in data[0]:
        return data[0]["generated_text"]
    return str(data)


# ============================
# JSON抽出（安定化）
# ============================
def extract_json(text: str):
    json_match = re.search(r'\{[\s\S]*\}', text)
    if not json_match:
        raise ValueError("AI返答にJSONが見つかりませんでした。")

    json_text = json_match.group(0)

    try:
        return json.loads(json_text)
    except json.JSONDecodeError:
        raise ValueError("抽出したJSONが壊れています。")


# ============================
# Streamlit UI
# ============================
st.set_page_config(page_title="勤務表自動生成（AI解析版）", layout="wide")
st.title("📘 勤務表自動生成システム（HuggingFaceAI連携）")
st.sidebar.header("Excelアップロード")

uploaded_file = st.sidebar.file_uploader("勤務表Excelをアップロード", type=["xlsx"])

if not uploaded_file:
    st.info("勤務表Excelをそのままアップロードしてください。")
    st.stop()


# ============================
# Excel → テキスト化（安全版）
# ============================
df_raw = pd.read_excel(uploaded_file, header=None)

if df_raw.shape[1] == 0:
    st.error("Excelに列がありません。主任の勤務表の形式を確認してください。")
    st.stop()

first_col = df_raw.iloc[:, 0]
text_data = "\n".join(first_col.fillna("").astype(str).tolist())

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
  - List of unique work codes such as 日1, 日2ホ, 公, 入浴, 有, 会議, etc.

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

        st.success("AIによるExcel解析が完了しました。")
        st.write("### 抽出された職員一覧（AI解析結果）")
        st.json(staff_list)

        st.write("### 抽出された勤務記号一覧（AI解析結果）")
        st.json(codes_list)

        # ============================
        # 職員ごとの設定UI
        # ============================
        staff_settings = {}
        staff_names = [s["name"] for s in staff_list if "name" in s]

        st.sidebar.markdown("## 職員ごとの設定")

        for name in staff_names:
            st.sidebar.markdown(f"#### {name}")

            universal = st.sidebar.checkbox(
                f"{name}: 万能枠",
                value=False,
                key=f"universal_{name}"
            )

            night_count = st.sidebar.number_input(
                f"{name}: 夜勤数",
                2, 6, 4,
                key=f"night_count_{name}"
            )

            night_double = st.sidebar.checkbox(
                f"{name}: 夜勤2連勤OK",
                value=True,
                key=f"night_double_{name}"
            )

            ng_list = st.sidebar.multiselect(
                f"{name}: NGペア（同じグループ勤務禁止）",
                staff_names,
                default=[],
                key=f"ng_{name}"
            )

            staff_settings[name] = {
                "universal": universal,
                "night_count": night_count,
                "night_double": night_double,
                "ng_pairs": ng_list,
            }

        # ============================
        # engine.py が必要とする空の DataFrame を作成
        # ============================
        df = pd.DataFrame(
            index=staff_names,
            columns=[f"{i+1}日" for i in range(31)]
        )

        # ============================
        # 最適化
        # ============================
        if st.button("勤務表を自動生成する"):
            with st.spinner("最適化エンジンが勤務表を計算中..."):
                result = solve_schedule(df, staff_settings)

                if result is None:
                    st.error("制約を満たす解が見つかりませんでした。条件を調整してください。")
                    st.stop()

                st.success("勤務表の自動生成が完了しました！")
                st.write("### 生成された勤務表")
                st.dataframe(result, use_container_width=True)

                excel_binary = export_excel(result)

                st.download_button(
                    "Excelファイルをダウンロード",
                    excel_binary,
                    "generated_schedule.xlsx"
                )
