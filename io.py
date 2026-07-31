# io.py
import pandas as pd
import io
from openpyxl import Workbook

# ============================
# Excel 読み込み
# ============================

def load_excel(file):
    """
    Excelファイルを読み込んで DataFrame を返す。
    index_col=0 で職員名を行インデックスにする。
    """
    df = pd.read_excel(file, index_col=0)
    return df


# ============================
# Excel 出力（勤務表を保存）
# ============================

def export_excel(df):
    """
    DataFrame を Excel バイナリに変換して返す。
    Streamlit の download_button に渡すために BytesIO を使う。
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer)
    return output.getvalue()
