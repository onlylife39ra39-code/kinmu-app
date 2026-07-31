# app.py
import streamlit as st
import pandas as pd
import json
import requests
import re

from io_utils import export_excel
from engine import solve_schedule


# ============================
# OpenRouter API
# ============================
OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"
OPENROUTER_MODEL = "meta-llama/llama-3.1-8b-instruct"
OPENROUTER_API_KEY = "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"


def call_ai(prompt: str) -> str:
    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json"
    }

    payload = {
        "model": OPENROUTER_MODEL,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.2,
        "max_tokens": 16000
    }

    response = requests.post(OPENROUTER_URL, headers=headers, json=payload)
    response.raise_for_status()
    data = response.json()
    return data["choices"][0]["message"]["content"]


# ============================
# JSON抽出
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
# グループ自動補正（セル結合対応）
# ============================
groups_full = []
current_group = None

for g in group_col:
    if "グループ" in g:
        current_group = g
    groups_full.append(current_group)


# ============================
# 役職自動補正（セル結合対応）
# ============================
roles_full = []
current_role = None

for r in role_col:
    if any(x in r for x in ["介護長", "介護主任", "長", "主任"]):
        current_role = r
    roles_full.append(current_role)


# ============================
# 職員名フィルタ（index同期）
# ============================
filtered_indices = []
filtered_names = []

for i, n in enumerate(name_col):
    if is_staff_name(n):
        filtered_indices.append(i)
        filtered_names.append(n.replace("☆", ""))

# index同期で役職・グループ抽出
filtered_roles = [roles_full[i] for i in filtered_indices]
filtered_groups = [groups_full[i] for i in filtered_indices]


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
# AI解析
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

        # index同期した役職・グループを上書き
        for idx, s in enumerate(staff_list):
            s["role"] = filtered_roles[idx]
            s["group"] = filtered_groups[idx]

        st.success("AIによるExcel解析が完了しました！")
        st.write("### 職員一覧（AI＋フィルタ同期＋自動補正）")
        st.json(staff_list)

        st.write("### 勤務記号一覧")
        st.json(codes_list)

        # ============================
        # 手動設定（session_stateで永続化）
        # ============================
        staff_settings = {}
        staff_names = [s["name"] for s in staff_list]

        st.sidebar.markdown("## 職員ごとの設定（永続化）")

        for idx, s in enumerate(staff_list):
            name = s["name"]

            # --- グループ ---
            if f"group_{name}" not in st.session_state:
                st.session_state[f"group_{name}"] = s["group"]

            group = st.sidebar.text_input(
                f"{name}: グループ",
                value=st.session_state[f"group_{name}"],
                key=f"group_input_{name}"
            )
            st.session_state[f"group_{name}"] = group

            # --- 役職 ---
            if f"role_{name}" not in st.session_state:
                st.session_state[f"role_{name}"] = s["role"]

            role = st.sidebar.text_input(
                f"{name}: 役職",
                value=st.session_state[f"role_{name}"],
                key=f"role_input_{name}"
            )
            st.session_state[f"role_{name}"] = role

            # --- その他設定 ---
            universal = st.sidebar.checkbox(
                f"{name}: 万能枠",
                value=st.session_state.get(f"universal_{name}", False),
                key=f"universal_{name}"
            )

            night_count = st.sidebar.number_input(
                f"{name}: 夜勤数",
                2, 6, 4,
                key=f"night_count_{name}"
            )

            night_double = st.sidebar.checkbox(
                f"{name}: 夜勤2連勤OK",
                value=st.session_state.get(f"night_double_{name}", True),
                key=f"night_double_{name}"
            )

            ng_list = st.sidebar.multiselect(
                f"{name}: NGペア",
                staff_names,
                default=st.session_state.get(f"ng_{name}", []),
                key=f"ng_{name}"
            )

            staff_settings[name] = {
                "role": role,
                "group": group,
                "universal": universal,
                "night_count": night_count,
                "night_double": night_double,
                "ng_pairs": ng_list,
            }

        # ============================
        # 空のDataFrame作成
        # ============================
        df = pd.DataFrame(
            index=staff_names,
            columns=[f"{i+1}日" for i in range(31)]
        )

        # ============================
        # 最適化エンジン
        # ============================
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
