# constraints.py
from ortools.sat.python import cp_model

# ============================
# v2.4 全制約セット
# ============================

def add_basic_constraints(model, x, staff_ids, num_days, shifts, work_shifts, is_night):
    """1人1日1勤務 / 月17勤務 / 夜勤明け明公 / 夜勤単発設定"""
    for s in staff_ids:
        # 1人1日1勤務
        for d in range(num_days):
            model.Add(sum(x[s, d, sh] for sh in shifts) == 1)

        # 月17勤務
        model.Add(sum(x[s, d, sh] for d in range(num_days) for sh in work_shifts) == 17)

        # 夜勤明けは明公
        for d in range(num_days - 1):
            model.Add(x[s, d+1, "明公"] == 1).OnlyEnforceIf(is_night[s, d])


def add_single_night_constraints(model, x, staff_ids, num_days, is_night, staff_settings):
    """夜勤単発設定（夜勤2連勤不可）"""
    for s in staff_ids:
        if not staff_settings[s].get("night_double", True):
            for d in range(num_days - 1):
                model.Add(
                    x[s, d+1, "夜1"] + x[s, d+1, "夜2"] == 0
                ).OnlyEnforceIf(is_night[s, d])


def add_daily_shift_constraints(model, x, staff_ids, num_days):
    """各勤務1日1人（公・明公・日1F以外）"""
    for d in range(num_days):
        for sh in ["早1", "早A", "早B", "早C", "日1", "日2ホ", "遅1A", "遅1B", "遅1C"]:
            model.Add(sum(x[s, d, sh] for s in staff_ids) == 1)

        # 夜勤セット
        model.Add(sum(x[s, d, "夜1"] for s in staff_ids) == 1)
        model.Add(sum(x[s, d, "夜2"] for s in staff_ids) == 1)


def add_ng_constraints(model, x, staff_ids, num_days, staff_settings):
    """人NG（同グループNG）"""
    for s in staff_ids:
        ng_list = staff_settings[s].get("ng_pairs", [])
        for ng_partner in ng_list:
            if ng_partner not in staff_ids:
                continue

            for d in range(num_days):
                # Aグループ
                model.Add(
                    x[s, d, "早A"] + x[s, d, "遅1A"] +
                    x[ng_partner, d, "早A"] + x[ng_partner, d, "遅1A"]
                    <= 1
                )

                # Bグループ
                model.Add(
                    x[s, d, "早B"] + x[s, d, "遅1B"] +
                    x[ng_partner, d, "早B"] + x[ng_partner, d, "遅1B"]
                    <= 1
                )

                # Cグループ
                model.Add(
                    x[s, d, "早C"] + x[s, d, "遅1C"] +
                    x[ng_partner, d, "早C"] + x[ng_partner, d, "遅1C"]
                    <= 1
                )


def add_bathing_ng(model, x, staff_ids, num_days, staff_settings):
    """入浴担当NG（早1・早A/B/C・日1）"""
    bathing_shifts = ["早1", "早A", "早B", "早C", "日1"]

    for s in staff_ids:
        ng_list = staff_settings[s].get("ng_pairs", [])
        for ng_partner in ng_list:
            if ng_partner not in staff_ids:
                continue

            for d in range(num_days):
                model.Add(
                    sum(x[s, d, sh] for sh in bathing_shifts) +
                    sum(x[ng_partner, d, sh] for sh in bathing_shifts)
                    <= 1
                )


def add_night_interval(model, x, staff_ids, num_days, is_night, staff_settings):
    """夜勤間隔（通常4日、万能枠は3日）"""
    for s in staff_ids:
        is_universal = staff_settings[s].get("universal", False)
        min_interval = 3 if is_universal else 4

        for d in range(num_days):
            for k in range(1, min_interval + 1):
                if d + k < num_days:
                    model.Add(
                        x[s, d+k, "夜1"] + x[s, d+k, "夜2"] == 0
                    ).OnlyEnforceIf(is_night[s, d])


def add_max_consecutive_work(model, x, staff_ids, num_days):
    """連勤最大3連勤（公・明公で途切れる）"""
    for s in staff_ids:
        for d in range(num_days - 3):
            # 4連勤禁止
            model.Add(
                sum(
                    1 - x[s, d+i, "公"] - x[s, d+i, "明公"]
                    for i in range(4)
                ) <= 3
            )


def add_penalties(model, x, staff_ids, num_days, staff_settings):
    """夜勤回数・万能枠ペナルティ"""
    penalties = []

    for s in staff_ids:
        is_universal = staff_settings[s].get("universal", False)
        target_night = staff_settings[s].get("night_count", 4)

        actual_night = sum(x[s, d, "夜1"] + x[s, d, "夜2"] for d in range(num_days))

        diff = model.NewIntVar(0, 10, f"diff_night_{s}")
        model.AddAbsEquality(diff, actual_night - target_night)

        weight = 10 if is_universal else 100
        penalties.append(diff * weight)

    model.Minimize(sum(penalties))


# ============================
# まとめて呼び出す関数
# ============================

def apply_all_constraints(model, x, staff_ids, num_days, shifts, work_shifts, is_night, staff_settings):
    add_basic_constraints(model, x, staff_ids, num_days, shifts, work_shifts, is_night)
    add_single_night_constraints(model, x, staff_ids, num_days, is_night, staff_settings)
    add_daily_shift_constraints(model, x, staff_ids, num_days)
    add_ng_constraints(model, x, staff_ids, num_days, staff_settings)
    add_bathing_ng(model, x, staff_ids, num_days, staff_settings)
    add_night_interval(model, x, staff_ids, num_days, is_night, staff_settings)
    add_max_consecutive_work(model, x, staff_ids, num_days)
    add_penalties(model, x, staff_ids, num_days, staff_settings)
