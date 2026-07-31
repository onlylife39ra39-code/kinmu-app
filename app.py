# app.py
import streamlit as st
import pandas as pd
import json
import re

from io_utils import export_excel
from engine import solve_schedule


# ============================
# AI呼び出し（ここだけ環境に合わせて書き換える）
# ============================
def call_ai(prompt: str) -> str:
    """
    ここにまーくんの使うAIの呼び出し処理を書く。
    例：OpenAI / Azure / Copilot Studio など。
    今はダミーとして、手動でJSONを返す形にしてある。
    """
    # ★★★ ここを実際のAI呼び出しに差し替える ★★★
    # 返すべきJSONの形：
    # {
    #   "staff": [
    #     {"name": "千明恵美", "role": "介護長", "group": null},
    #     {"name": "浦野裕太", "role": "介護主任", "group": null},
    #     {"name": "茂木最恵", "role": null, "group": "Aグループ"},
    #     ...
    #   ],
    #   "codes": ["日1", "公", "入浴", ...]
    # }
    dummy = {
        "staff": [],
        "codes": []
    }
    return json.dumps(dummy, ensure_ascii=False)


# ============================
# Streamlit UI 設定
# ============================
st.set_page_config(page_title="勤務表自動生成（全部AI解析版）", layout="wide")

st.title("📘 勤務表自動生成システム（Excel全部AI解析版）")
st.sidebar.header("Excelアップロード")


# ============================
# Excel アップロード
# ============================
uploaded_file = st.sidebar.file_uploader("主任作成の勤務表Excelをアップロード", type=["xlsx"])

if not uploaded_file:
    st.info("主任が作った勤務表Excelをそのままアップロードしてね。")
    st.stop()

# ============================
# Excel → テキスト化（AIに渡す前処理）
# ============================
df_raw = pd.read_excel(uploaded_file, header=None)

# ここでは 1列目を対象にしているが、必要なら複数列を結合してもOK
text_data = "\n".join(df_raw[0].astype(str).tolist())

st.write("### 読み込みデータ（1列目のテキスト化プレビュー）")
st.text_area("Excel 1列目テキスト", text_data, height=200)


# ============================
# AI に構造化を依頼
# ============================
prompt = f"""
以下は介護施設の勤務表Excelの1列目です。
このデータを構造化してください。

抽出したい項目：
- 職員名（記号は除外）
- 役職名（介護長、主任など）
- グループ名（Aグループなど）
- 勤務記号（日1、公、入浴、有、会議など）
- 数字（勤務数）
- その他のメタ情報

出力形式（JSON）：
{{
  "staff": [
    {{"name": "千明恵美", "role": "介護長", "group": null}},
    {{"name": "浦野裕太", "role": "介護主任", "group": null}},
    {{"name": "茂木最恵", "role": null, "group": "Aグループ"}},
    ...
  ],
  "codes": ["日1", "公", "入浴", ...]
}}

データ:
{text_data}
"""

st.write("### AI に渡すプロンプト（確認用）")
with st.expander("プロンプトを見る"):
    st.code(prompt)

if st.button("AIでExcelを解析して勤務表を構造化する"):
    with st.spinner("AIがExcelを解析中..."):
        ai_response = call_ai(prompt)

        try:
            parsed = json.loads(ai_response)
        except json.JSONDecodeError:
            st.error("AIの返答がJSONとして読み込めませんでした。プロンプトかAI側の設定を確認してください。")
            st.stop()

        staff_list = parsed.get("staff", [])
        codes_list = parsed.get("codes", [])

        if not staff_list:
            st.error("AIから職員情報が取得できませんでした。プロンプトを調整する必要があります。")
            st.stop()

        st.success("AIによるExcel解析が完了しました。")
        st.write("### 抽出された職員一覧（AI解析結果）")
        st.json(staff_list)

        st.write("### 抽出された勤務記号一覧（AI解析結果）")
        st.json(codes_list)

        # ============================
        # 職員ごとの設定UI（AI抽出結果を元に）
        # ============================
        staff_settings = {}

        st.sidebar.markdown("## 職員ごとの設定（AI抽出済み職員）")

        # 名前リストだけ抜き出し
        staff_names = [s["name"] for s in staff_list if "name" in s]

        for name in staff_names:
            st.sidebar.markdown(f"#### {name} の設定")

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
        # 最適化ボタン
        # ============================
        if st.button("AI解析結果を使って勤務表を自動生成する"):
            with st.spinner("最適化エンジンが勤務表を計算中..."):
                # solve_schedule は AIからの staff_list と UI設定を受け取る想定
                result = solve_schedule(staff_list, staff_settings)

                if result is not None:
                    st.success("勤務表の自動生成が完了しました！")
                    st.write("### 生成された勤務表（編集不可）")
                    st.dataframe(result, use_container_width=True)

                    excel_binary = export_excel(result)

                    st.download_button(
                        "Excelファイルをダウンロード",
                        excel_binary,
                        "generated_schedule.xlsx"
                    )
                else:
                    st.error("制約を満たす解が見つかりませんでした。条件を緩和してください。")
