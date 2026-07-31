# app.py
import streamlit as st
import pandas as pd
import json
import requests
import re
import os

from io_utils import export_excel
from engine import solve_schedule

# ============================
# HuggingFace API（無料・カード不要）
# ============================
HF_URL = "https://api-inference.huggingface.co/models/mistralai/Mistral-7B-Instruct-v0.2"

def call_ai(prompt: str) -> str:
    HF_API_KEY = os.getenv("HF_API_KEY")

    if not HF_API_KEY:
        st.error("HuggingFace APIキーが読み込めていません（Secretsに設定してください）")
        return None

    headers = {
        "Authorization": f"Bearer {HF_API_KEY}",
        "Content-Type": "application/json"
    }

    payload = {
        "inputs": prompt
    }

    resp = requests.post(HF_URL, headers=headers, json=payload)

    if resp.status_code != 200:
        st.error(f"HuggingFace API Error: {resp.status_code}")
        st.code(resp.text)
        return None

    try:
        return resp.json()[0]["generated_text"]
    except Exception:
        st.error("AI返答の解析に失敗しました（JSON形式が不正）")
        st.code(resp.text)
        return None


# ============================
# JSON抽出
# ============================
def extract_json(text: str):
    m = re.search(r'\{[\s\S]*\}', text)
    if not m:
        raise ValueError("AI返答にJSONが見つかりませんでした。")
    return json.loads(m.group(0))


# ============================
# 職員名フィルタ
# ============================
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


# ============================
# 役職・グループを「上にさかのぼって」取得
# ============================
def find_last_role_above(index, role_col):
    for i in range(index, -1, -1):
        r = role_col[i]
        if any(x in r for x in ["介護長", "介護主任", "長", "主任"]):
            return r
    return None

def find_last_group_above(index, group_col):
    for i in range(index, -1, -1):
        g = group_col[i]
        if "グループ" in g:
            return g
    return None


# ============================
# 永続保存 JSON
# ============================
SETTINGS_FILE = "staff_settings.json"

def load_staff_settings():
    if os.path.exists(SETTINGS_FILE):
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_staff_settings(settings: dict):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)


# ============================
# Streamlit UI
# ============================
st.set_page_config(page_title="勤務表自動生成（AI勤務エンジン）", layout="wide")
st.title("📘 勤務表自動生成（AI勤務エンジン）")

st.sidebar.header("Excelアップロード")
uploaded_file = st.sidebar.file_uploader("勤務表Excelをアップロード", type=["xlsx"])

if not uploaded_file:
    st.info("勤務表Excelをアップロードしてください。")
    st.stop()


# ============================
# Excel読み込み
# ============================
df_raw = pd.read_excel(uploaded_file, header=None)

name_col = df_raw.iloc[:, 1].fillna("").astype(str).tolist()
group_col = df_raw.iloc[:, 0].fillna("").astype(str).tolist()
role_col = df_raw.iloc[:, 0].fillna("").astype(str).tolist()


# ============================
# 職員名フィルタ & index同期
# ============================
filtered_indices = []
filtered_names = []

for i, n in enumerate(name_col):
    if is_staff_name(n):
        filtered_indices.append(i)
        filtered_names.append(n.replace("☆", ""))

text_data = "\n".join(filtered_names)

filtered_roles = [find_last_role_above(i, role_col) for i in filtered_indices]
filtered_groups = [find_last_group_above(i, group_col) for i in filtered_indices]

st.write("### 抽出された職員名（フィルタ後）")
st.json(filtered_names)


# ============================
# AIプロンプト
# ============================
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

データ:
{text_data}
"""

with st.expander("AIに渡すプロンプト（確認用）"):
    st.code(prompt)


# ============================
# 永続設定読み込み
# ============================
persistent_settings = load_staff_settings()


# ============================
# AI解析
# ============================
if st.button("AIでExcelを解析する"):
    with st.spinner("AIがExcelを解析中..."):
        ai_response = call_ai(prompt)

        if ai_response is None:
            st.stop()

        try:
            parsed = extract_json(ai_response)
        except Exception as e:
            st.error(f"AI返答の解析に失敗しました: {e}")
            st.text(ai_response)
            st.stop()

        staff_list = parsed.get("staff", [])
        codes_list = parsed.get("codes", [])

        for idx, s in enumerate(staff_list):
            s["role"] = filtered_roles[idx]
            s["group"] = filtered_groups[idx]

        st.success("AIによるExcel解析が完了しました！")
        st.write("### 職員一覧（AI＋役職・グループ補正）")
        st.json(staff_list)

        st.write("### 勤務記号一覧")
        st.json(codes_list)

        staff_settings = persistent_settings.copy()
        staff_names = [s["name"] for s in staff_list]

        st.sidebar.markdown("## 職員ごとの設定（永続保存）")

        delete_targets = []

        for idx, s in enumerate(staff_list):
            name = s["name"]

            base = staff_settings.get(name, {})
            base_role = base.get("role", s["role"])
            base_group = base.get("group", s["group"])
            base_universal = base.get("universal", False)
            base_night_count = base.get("night_count", 4)
            base_night_double = base.get("night_double", True)
            base_ng_pairs = base.get("ng_pairs", [])

            st.sidebar.markdown(f"### {name}")

            group = st.sidebar.text_input(
                f"{name}: グループ",
                value=base_group if base_group else "",
                key=f"group_input_{name}"
            )

            role = st.sidebar.text_input(
                f"{name}: 役職",
                value=base_role if base_role else "",
                key=f"role_input_{name}"
            )

            universal = st.sidebar.checkbox(
                f"{name}: 万能枠",
                value=base_universal,
                key=f"universal_{name}"
            )

            night_count = st.sidebar.number_input(
                f"{name}: 夜勤数",
                2, 6, int(base_night_count),
                key=f"night_count_{name}"
            )

            night_double = st.sidebar.checkbox(
                f"{name}: 夜勤2連勤OK",
                value=base_night_double,
                key=f"night_double_{name}"
            )

            ng_list = st.sidebar.multiselect(
                f"{name}: NGペア",
                staff_names,
                default=base_ng_pairs,
                key=f"ng_{name}"
            )

            delete_flag = st.sidebar.checkbox(
                f"{name}: この職員を削除（退職など）",
                value=False,
                key=f"delete_{name}"
            )
            if delete_flag:
                delete_targets.append(name)

            if name not in delete_targets:
                staff_settings[name] = {
                    "role": role if role else None,
                    "group": group if group else None,
                    "universal": universal,
                    "night_count": int(night_count),
                    "night_double": night_double,
                    "ng_pairs": ng_list,
                }

        for name in delete_targets:
            if name in staff_settings:
                del staff_settings[name]

        save_staff_settings(staff_settings)
        st.success("職員設定を保存しました（永続保存）。")

        active_staff_names = list(staff_settings.keys())
        df = pd.DataFrame(
            index=active_staff_names,
            columns=[f"{i+1}日" for i in range(31)]
        )

        if st.button("勤務表を自動生成する"):
            with st.spinner("最適化エンジンが勤務表を計算中..."):
                result = solve_schedule(df, staff_settings)

                if result is None:
                    st.error("制約を満たす解が見つかりませんでした。条件を調整してください。")
                    st.stop()

                st.success("勤務表の自動生成が完了しました！")
                st.dataframe(result, use_container_width=True)

                excel_binary = export_excel(result)
                st.download_button(
                    "Excelファイルをダウンロード",
                    excel_binary,
                    "generated_schedule.xlsx"
                )
