import streamlit as st
import pandas as pd
import numpy as np
from ortools.sat.python import cp_model
import io

st.set_page_config(page_title="勤務表自動生成 v2.4", layout="wide")

st.title("📘 勤務表自動生成システム v2.4")
st.sidebar.header("設定アップロード")

# --- ダミーデータ作成機能 (テスト用) ---
def create_sample_data():
    names = [f"職員{i}" for i in range(1, 11)]
    df = pd.DataFrame(index=names, columns=[f"{i}日" for i in range(1, 31)])
    return df

# --- 最適化エンジン ---
def solve_schedule(df, staff_settings):
    num_days = len(df.columns)
    staff_ids = df.index.tolist()
    
    model = cp_model.CpModel()
    
    # 勤務記号定義
    shifts = ["早1", "早A", "早B", "早C", "日1", "日2ホ", "遅1A", "遅1B", "遅1C", "夜1", "夜2", "公", "明公", "日1F"]
    work_shifts = [s for s in shifts if s not in ["公", "明公"]]
    
    # 変数作成: x[staff, day, shift]
    x = {}
    for s in staff_ids:
        for d in range(num_days):
            for sh in shifts:
                x[s, d, sh] = model.NewBoolVar(f'x_{s}_{d}_{sh}')

    # --- 制約実装 ---
    for s in staff_ids:
        setting = staff_settings[s]
        # 1. 1人1日1勤務
        for d in range(num_days):
            model.Add(sum(x[s, d, sh] for sh in shifts) == 1)
        
        # 2. 月17日勤務
        model.Add(sum(x[s, d, sh] for d in range(num_days) for sh in work_shifts) == 17)

        # 3. 夜勤の後の明公
        for d in range(num_days - 1):
            is_night = model.NewBoolVar(f'is_night_{s}_{d}')
            model.AddMaxEquality(is_night, [x[s, d, "夜1"], x[s, d, "夜2"]])
            model.Add(x[s, d+1, "明公"] == 1).OnlyEnforceIf(is_night)

    # 4. 各勤務に1日1人 (日1F, 公, 明公以外)
    for d in range(num_days):
        for sh in ["早1", "早A", "早B", "早C", "日1", "日2ホ", "遅1A", "遅1B", "遅1C"]:
            model.Add(sum(x[s, d, sh] for s in staff_ids) == 1)
        # 夜勤セット
        model.Add(sum(x[s, d, "夜1"] for s in staff_ids) == 1)
        model.Add(sum(x[s, d, "夜2"] for s in staff_ids) == 1)

    # 5. 万能枠・連勤・夜勤回数の最適化スコア
    penalties = []
    for s in staff_ids:
        is_universal = staff_settings[s].get('universal', False)
        # 夜勤回数
        target_night = staff_settings[s].get('night_count', 4)
        actual_night = sum(x[s, d, "夜1"] + x[s, d, "夜2"] for d in range(num_days))
        
        diff = model.NewIntVar(0, 10, '')
        model.AddAbsEquality(diff, actual_night - target_night)
        
        weight = 10 if is_universal else 100
        penalties.append(diff * weight)

    model.Minimize(sum(penalties))

    # 解く
    solver = cp_model.CpSolver()
    status = solver.Solve(model)

    if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
        # 結果をDataFrameに変換
        res_df = df.copy()
        for s in staff_ids:
            for d in range(num_days):
                for sh in shifts:
                    if solver.Value(x[s, d, sh]):
                        res_df.loc[s, f"{d+1}日"] = sh
        return res_df
    else:
        return None

# --- UI 部分 ---
uploaded_file = st.sidebar.file_uploader("勤務表Excelをアップロード", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, index_col=0)
    st.write("### 読み込みデータ", df)
    
    # 職員設定（デモ用。実際はExcelやサイドバーから取得）
    staff_settings = {}
    for name in df.index:
        staff_settings[name] = {
            'universal': st.sidebar.checkbox(f"{name}: 万能枠", value=False),
            'night_count': st.sidebar.number_input(f"{name}: 夜勤数", 2, 6, 4)
        }

    if st.button("勤務表を自動生成する"):
        with st.spinner("最適化計算中..."):
            result = solve_schedule(df, staff_settings)
            
            if result is not None:
                st.success("生成完了！")
                st.write(result)
                
                # Excelダウンロード
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    result.to_excel(writer)
                st.download_button("Excelファイルをダウンロード", output.getvalue(), "generated_schedule.xlsx")
            else:
                st.error("制約を満たす解が見つかりませんでした。条件を緩和してください。")
else:
    st.info("左側のサイドバーから既存のExcel勤務表をアップロードしてください。")
