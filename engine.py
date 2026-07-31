# engine.py
import pandas as pd
from ortools.sat.python import cp_model
from constraints import apply_all_constraints

# ============================
# 最適化エンジン（v2.4）
# ============================

def solve_schedule(df, staff_settings):
    num_days = len(df.columns)
    staff_ids = df.index.tolist()

    model = cp_model.CpModel()

    # 勤務記号定義
    shifts = [
        "早1", "早A", "早B", "早C",
        "日1", "日2ホ",
        "遅1A", "遅1B", "遅1C",
        "夜1", "夜2",
        "公", "明公",
        "日1F",
    ]
    work_shifts = [s for s in shifts if s not in ["公", "明公"]]

    # 変数作成: x[staff, day, shift]
    x = {}
    for s in staff_ids:
        for d in range(num_days):
            for sh in shifts:
                x[s, d, sh] = model.NewBoolVar(f"x_{s}_{d}_{sh}")

    # 夜勤判定用変数
    is_night = {}
    for s in staff_ids:
        for d in range(num_days):
            is_night[s, d] = model.NewBoolVar(f"is_night_{s}_{d}")
            model.AddMaxEquality(is_night[s, d], [x[s, d, "夜1"], x[s, d, "夜2"]])

    # ============================
    # v2.4 全制約を適用
    # ============================
    apply_all_constraints(
        model=model,
        x=x,
        staff_ids=staff_ids,
        num_days=num_days,
        shifts=shifts,
        work_shifts=work_shifts,
        is_night=is_night,
        staff_settings=staff_settings
    )

    # ============================
    # ソルバー実行
    # ============================
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 20  # 安定化のため

    status = solver.Solve(model)

    # ============================
    # 結果を DataFrame に変換
    # ============================
    if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        result_df = df.copy()

        for s in staff_ids:
            for d in range(num_days):
                for sh in shifts:
                    if solver.Value(x[s, d, sh]):
                        result_df.loc[s, f"{d+1}日"] = sh

        return result_df

    return None
