import streamlit as st
import pandas as pd
import json
import re

st.set_page_config(page_title="勤務表AI（Google AIモード自動化版）", layout="wide")

st.title("📘 勤務表AI（Google AIモード自動化版）")
st.write("Google AI モードを使って完全無料で勤務表解析できます。")

# -----------------------------
# Excel アップロード
# -----------------------------
st.sidebar.header("Excelアップロード")
uploaded_file = st.sidebar.file_uploader("勤務表Excelをアップロード", type=["xlsx"])

if not uploaded_file:
    st.info("勤務表Excelをアップロードしてください。")
    st.stop()

df_raw = pd.read_excel(uploaded_file, header=None)

name_col = df_raw.iloc[:, 1].fillna("").astype(str).tolist()
group_col = df_raw.iloc[:, 0].fillna("").astype(str).tolist()
role_col = df_raw.iloc[:, 0].fillna("").astype(str).tolist()

# -----------------------------
# 職員名フィルタ
# -----------------------------
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

filtered_names = [n.replace("☆", "") for n in name_col if is_staff_name(n)]
text_data = "\n".join(filtered_names)

st.write("### 抽出された職員名")
st.json(filtered_names)

# -----------------------------
# Google AI モードに貼るプロンプト自動生成
# -----------------------------
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
- JSONのトップレベルキーは必ず "staff" と "codes" の2つにしてください。

データ:
{text_data}
"""

st.write("### ▼ Google AI モードに貼るテキスト（自動生成）")
st.code(prompt)

st.write("### ▼ Google AI モード（Gemini）をここで直接使えます")
st.components.v1.iframe("https://aistudio.google.com/app/prompts/new_chat", height=700)

# -----------------------------
# JSON貼り付け欄（自動整形付き）
# -----------------------------
st.write("### ▼ Google AI モードの返す JSON を貼り付けてください")
raw_json = st.text_area("AIから返ってきたJSONを貼り付ける欄", height=300)

def extract_json(text: str):
    m = re.search(r'\{[\s\S]*\}', text)
    if not m:
        raise ValueError("JSONが見つかりませんでした。")
    return json.loads(m.group(0))

if st.button("JSONを自動整形する"):
    try:
        parsed = extract_json(raw_json)
        st.success("JSONを自動整形しました！")
        st.json(parsed)
    except Exception as e:
        st.error(f"JSON解析に失敗しました: {e}")

# -----------------------------
# 勤務表生成
# -----------------------------
if st.button("勤務表を生成する"):
    try:
        parsed = extract_json(raw_json)
    except Exception as e:
        st.error(f"JSON解析に失敗しました: {e}")
        st.stop()

    st.success("JSON解析が完了しました！")
    st.write("### 職員一覧")
    st.json(parsed.get("staff", []))

    st.write("### 勤務記号一覧")
    st.json(parsed.get("codes", []))

    st.write("### ▼ ここから勤務表生成ロジックを追加できます（まーくんの既存コードを統合可能）")
