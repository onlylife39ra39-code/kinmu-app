# app.py
import streamlit as st
import pandas as pd
import json
import requests
import re

from io_utils import export_excel
from engine import solve_schedule


# ============================
# OpenRouter API（Cloudで安定）
# ============================
OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"
OPENROUTER_MODEL = "meta-llama/llama-3.1-8b-instruct"  # 日本語強い高速モデル
OPENROUTER_API_KEY = "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"


def call_ai(prompt: str) -> str:
    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json"
    }

    payload = {
        "model": OPENROUTER_MODEL,
        "messages": [
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.2,
        "max_tokens": 16000  # 長いJSONでも途切れないように多め
    }

    response = requests.post(OPENROUTER_URL, headers=headers, json=payload)
    response.raise_for_status()
    data = response.json()

    return data["choices"][0]["message"]["content"]


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
# 職員名フィルタ（まーくん勤務表専用）
# ============================
def is_staff_name(text: str) -> bool:
    if not text:
        return False

    t = text.replace("☆", "").strip()

    # 除外ワード
    if t in ["月間予定"]:
        return False

    # 勤務時間の行を除外
    if any(x in t for x in ["～", ":", "勤務時間", "週", "月～金", "土日祝"]):
        return False

    # グループ名・役職名を除外
    if any(x in t for x in ["グループ", "介護長", "介護主任", "新入職員", "パート"]):
        return False

    # 漢字・ひらがな・カタカナのみ
    if re.match(r'^[\u4E00-\u9FFF\u3040-\u309F\u30A0-\u30FF]+$', t):
        return True

    return False


# ============================
# Streamlit UI
# ============================
st.set_page_config(page_title="勤務表自動生成（OpenRouter AI版）", layout="wide")
st.title("📘 勤務表自動生成システム（OpenRouter AI連携）")
st.sidebar.header("Excelアップロード")

uploaded_file = st.sidebar.file_uploader("勤務表Excelをアップロード", type=["xlsx"])

if not uploaded_file:
    st.info("勤務表Excelをそのままアップロードしてください。")
    st.stop()


# ============================
# Excel → 職員名抽出（2列目＋フィルタ）
# ============================
df_raw = pd.read_excel(uploaded_file, header=None)

if df_raw.shape[1] < 2:
    st.error("Excelの2列目に職員名がありません。主任の勤務表の形式を確認してください。")
    st.stop()

name_col = df_raw.iloc[:, 1]
raw_names = name_col.fillna("").astype(str).tolist()

filtered_names = [n.replace("☆", "") for n in raw_names if is_staff_name(n)]

st.write("### 抽出された職員名（フィルタ後）")
st.json(filtered_names)

text_data = "\n".join(filtered_names)


# ============================
# AI 解析プロンプト（JSONのみ返させる版）
# ============================
prompt = f"""
あなたは介護施設の勤務表Excelを解析するAIです。

以下は主任が作成した勤務表Excelの「職員名が縦に並んだ列」です。
このデータから次を抽出してください：

1. staff（職員一覧）
  - name（職員名）
  - role（役職：介護長・介護主任など）
  - group（Aグループ・Bグループなど）

2. codes（勤務記号一覧）
  - 日1, 日2ホ, 公, 入浴, 有, 会議 など

重要：
- 出力は JSON のみとし、説明文・Pythonコード・文章は一切含めないでください。
- 必ず次の形式で返してください（キー名・構造を厳守）：

{{
  "staff": [
    {{ "name": "千明恵美", "role": "介護長", "group": null }},
    {{ "name": "浦野裕太", "role": "介護主任", "group": null }},
    {{ "name": "茂木最恵", "role": null, "group": "Aグループ" }}
  ],
  "codes": ["日1", "公", "入浴"]
}}

上記はあくまで例です。実際のデータに基づいて staff と codes を生成してください。

データ:
{text_data}
"""

st.write("### AIに渡すプロンプト（確認用）")
with st.expander("プロンプトを見る"):
    st.code(prompt)


# ============================
# AI 解析実行（OpenRouter）
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

        st.success("AIによるExcel解析が完了しました！")
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
