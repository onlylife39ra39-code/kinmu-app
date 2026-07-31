# app.py
import streamlit as st
import pandas as pd

# io.py → io_utils.py に変更した前提
from io_utils import load_excel, export_excel

# 最適化エンジン
from engine import solve_schedule


st.set_page_config(page_title="勤務表自動生成 v2.4", layout="wide")

st.title("📘 勤務表自動生成システム v2.4")
st.sidebar.header("設定アップロード")

uploaded_file = st.sidebar.file_uploader("勤務表Excelをアップロード", type=["xlsx"])

if uploaded_file:
    df = load_excel(uploaded_file)

    st.write("### 読み込みデータ（編集不可）")
    st.dataframe(df, use_container_width=True)

    staff_settings = {}

    for name in df.index:
        st.sidebar.markdown(f"#### {name} の設定")

        universal = st.sidebar.checkbox(f"{name}: 万能枠", value=False)
        night_count = st.sidebar.number_input(f"{name}: 夜勤数", 2, 6, 4)
        night_double = st.sidebar.checkbox(f"{name}: 夜勤2連勤OK", value=True)

        ng_list = st.sidebar.multiselect(
            f"{name}: NGペア（同じグループ勤務禁止）",
            df.index.tolist(),
            default=[]
        )

        staff_settings[name] = {
            "universal": universal,
            "night_count": night_count,
            "night_double": night_double,
            "ng_pairs": ng_list,
        }

    if st.button("勤務表を自動生成する"):
        with st.spinner("最適化計算中..."):
            result = solve_schedule(df, staff_settings)

            if result is not None:
                st.success("生成完了！")
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

else:
    st.info("左側のサイドバーから既存のExcel勤務表をアップロードしてください。")
