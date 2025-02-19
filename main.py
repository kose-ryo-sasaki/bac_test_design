import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.colors as mcolors
import seaborn as sns
import openpyxl
from openpyxl.styles import PatternFill, Border, Side
from io import BytesIO
from datetime import datetime

# 🖥 ページのレイアウトをワイドに設定
st.set_page_config(layout="wide")

# **9×12の表の状態を保持**
if "reshaped_df" not in st.session_state:
    st.session_state["reshaped_df"] = None
if "color_data" not in st.session_state:
    st.session_state["color_data"] = None
if "color_map" not in st.session_state:
    st.session_state["color_map"] = None
if "row_labels" not in st.session_state:
    st.session_state["row_labels"] = None
if "file_name" not in st.session_state:
    st.session_state["file_name"] = None

# タイトル
st.title("Sample Matrix Sheet Creator")

# CSVファイルのアップロード
uploaded_file = st.file_uploader("CSVファイルをアップロードしてください", type=["csv"])

if uploaded_file is not None:
    # CSVをDataFrameに読み込む
    df = pd.read_csv(uploaded_file)
    st.session_state["file_name"] = uploaded_file.name  # アップロードしたファイル名

    # **ユーザーが編集可能なデータフレーム**
    edited_df = st.data_editor(df, num_rows="dynamic")

    # **変更を適用するボタン**
    if st.button("変更を適用"):
        df = edited_df.copy()  # 編集内容を反映

        # 9×12の枠を準備
        table_data = [[""] * 12 for _ in range(9)]
        color_data = [[""] * 12 for _ in range(9)]  # 色情報

        row_idx, col_idx = 0, 0  # 9×12の表の配置用

        # **データを処理**
        for _, row in df.iterrows():
            sample_name = row["sample_name"]
            conc = row["conc"]
            iter_size = row["iter"]  # iter のサイズ（試験回数）
            bac = row["bac"]  # `bac` の値
            newline_flag = row["newline_flag"]  # 改行フラグ
            blank_flag = row["blank_flag"]  # 空白フラグ

            # **データをセルに詰める**
            for i in range(1, iter_size + 1):
                if col_idx >= 12:  # 12列を超えたら次の行へ
                    row_idx += 1
                    col_idx = 0
                    if row_idx >= 9:  # 9行を超えたら終了
                        break

                table_data[row_idx][col_idx] = f"{sample_name}_{conc}_{i}"
                color_data[row_idx][col_idx] = bac  # `bac` を色識別用に格納
                col_idx += 1  # 次のセルへ移動

            # **空白フラグがある場合、その `iter` の後に空白セルを挿入**
            for _ in range(blank_flag):
                if col_idx >= 12:  # 12列を超えたら次の行へ
                    row_idx += 1
                    col_idx = 0
                    if row_idx >= 9:  # 9行を超えたら終了
                        break
                table_data[row_idx][col_idx] = ""  # 空白セルを追加
                col_idx += 1

            # **改行フラグが立っている場合は改行**
            if newline_flag == 1:
                row_idx += 1
                col_idx = 0

            # **9行を超えたら終了**
            if row_idx >= 9:
                break

        # **9×12の表の行・列のインデックスを変更**
        row_labels = list("ABCDEFGHI")[:9]  # A, B, C, D, ...
        col_labels = list(range(1, 13))  # 1, 2, 3, ..., 12
        reshaped_df = pd.DataFrame(table_data, index=row_labels, columns=col_labels)

        # **色のマッピングを定義**
        pastel_palette = sns.color_palette("pastel", len(df["bac"].unique()))  # `seaborn` のパステルカラー
        unique_bac = df["bac"].unique()
        color_map = {bac: mcolors.to_hex(pastel_palette[i]) for i, bac in enumerate(unique_bac)}

        # **状態を保持**
        st.session_state["reshaped_df"] = reshaped_df
        st.session_state["color_data"] = color_data
        st.session_state["color_map"] = color_map
        st.session_state["row_labels"] = row_labels

# **9×12の表が存在する場合のみ表示**
if st.session_state["reshaped_df"] is not None:
    st.write("### 試験データ表")

    # **系列（凡例）を表示**
    legend_html = "".join([f'<div style="display: inline-block; width: 20px; height: 20px; background-color: {st.session_state["color_map"][bac]}; margin-right: 5px;"></div> {bac}' for bac in st.session_state["color_map"].keys()])
    st.markdown(legend_html, unsafe_allow_html=True)

    # **表の表示（色付き）**
    styled_df = st.session_state["reshaped_df"].style.apply(
        lambda x: [
            f"background-color: {st.session_state['color_map'].get(st.session_state['color_data'][st.session_state['row_labels'].index(x.name)][col], '')}; color: black;"
            for col in range(len(x))
        ],
        axis=1
    )
    st.table(styled_df)

# **CSVファイルを作成**
def create_csv_file(df):
    output = BytesIO()
    df.to_csv(output, index=False, encoding="utf-8-sig")  # **CSVファイルをUTF-8（BOM付き）で保存**
    output.seek(0)  # **バッファの先頭に戻る**
    return output

# **Excelファイルを作成**
def create_excel_file(df, color_data, color_map, file_name):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        wb = writer.book
        ws = wb.create_sheet(title="9x12_Table")

        # **1行目にファイル名と日付**
        ws["A1"] = f"File: {file_name}"
        ws["B1"] = f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"

        # **B2セルから列名 (1,2,3,...) を追加**
        for j, col_name in enumerate(df.columns, start=2):  # B2, C2, D2...
            ws.cell(row=2, column=j, value=col_name)

        # **A3セルから行名 (A,B,C,...) を追加**
        for i, row_name in enumerate(df.index, start=3):  # A3, A4, A5...
            ws.cell(row=i, column=1, value=row_name)

        # **B3セルからデータを入力**
        for i, row in enumerate(df.index):
            for j, col in enumerate(df.columns):
                cell = ws.cell(row=i + 3, column=j + 2, value=df.loc[row, col])  # データはB3から開始
                bac_value = color_data[i][j]

                # **色を適用**
                if bac_value in color_map:
                    cell.fill = PatternFill(start_color=color_map[bac_value][1:], end_color=color_map[bac_value][1:], fill_type="solid")

                # **罫線を適用**
                thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
                cell.border = thin_border

        # **右側に系列情報を表示**
        ws["O2"] = "Legend"
        for i, (bac, color) in enumerate(color_map.items()):
            ws.cell(row=i + 3, column=15, value=bac)
            ws.cell(row=i + 3, column=16).fill = PatternFill(start_color=color[1:], end_color=color[1:], fill_type="solid")

    processed_data = output.getvalue()
    return processed_data



# **CSVダウンロードボタン**
if uploaded_file is not None:
    st.download_button(
        label="修正後の元データをダウンロード",
        data=create_csv_file(edited_df),  # **修正後のデータをCSVとしてダウンロード**
        file_name="edited_data.csv",
        mime="text/csv"
    )


# **ダウンロードボタン**
if st.session_state["reshaped_df"] is not None and st.session_state["file_name"] is not None:
    st.download_button(
        label="マトリクスをダウンロード",
        data=create_excel_file(
            st.session_state["reshaped_df"],
            st.session_state["color_data"],
            st.session_state["color_map"],  # ここで `color_map` を `session_state` から取得
            st.session_state["file_name"]
        ),
        file_name="9x12_table.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

