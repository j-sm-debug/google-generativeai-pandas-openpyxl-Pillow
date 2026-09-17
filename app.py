import streamlit as st
import google.generativeai as genai
import PIL.Image
import pandas as pd
import io
import json
import os
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# --- ページ設定 ---
st.set_page_config(page_title="売上日計表 自動照合システム", layout="wide")

# --- ログイン認証処理 ---
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if st.session_state["password_correct"]:
        return True

    st.title("🔒 社内関係者専用ログイン")
    password_input = st.text_input("パスワードを入力してください", type="password")
    
    if st.button("ログイン"):
        if password_input == st.secrets.get("PASSWORD"):
            st.session_state["password_correct"] = True
            st.rerun()
        else:
            st.error("パスワードが正しくありません。")
    
    return False

if not check_password():
    st.stop()

# --- メイン画面 ---
st.title("📊 売上日計表 自動照合システム")
st.markdown("店舗からの売上日計表画像を自動照合し、全店舗マスタ（店舗.xlsx）と比較して**未提出店舗の抽出**およびExcel報告書を出力します。")

# SecretsからAPIキーを取得
api_key = st.secrets.get("GEMINI_API_KEY")

# 店舗マスタの読み込み
st.subheader("1. 店舗マスタ（調査対象店舗リスト）")
master_file = st.file_uploader("店舗マスタ（店舗.xlsx）を更新する場合はアップロードしてください（任意）", type=['xlsx', 'xls'])

master_df = None
if master_file:
    master_df = pd.read_excel(master_file)
    st.info("Uploaded: カスタム店舗マスタを使用します。")
elif os.path.exists("店舗.xlsx"):
    master_df = pd.read_excel("店舗.xlsx")
    st.success("✅ リポジトリ内の `店舗.xlsx`（116店舗）を自動読み込みしました。")
else:
    st.warning("⚠️ 店舗マスタ（店舗.xlsx）が読み込まれていません。未提出店舗の抽出を行う場合は画像をアップロードするかGitHubへ店舗.xlsxを追加してください。")

st.subheader("2. 日計表画像のアップロード")
uploaded_files = st.file_uploader(
    "売上日計表の画像をアップロードしてください（複数選択可・ドラッグ＆ドロップ対応）", 
    type=['png', 'jpg', 'jpeg'], 
    accept_multiple_files=True
)

if st.button("照合と未提出確認を開始する", type="primary"):
    if not api_key:
        st.error("システムエラー: APIキーが設定されていません。StreamlitのSecretsを設定してください。")
    elif not uploaded_files:
        st.warning("画像を1枚以上アップロードしてください。")
    else:
        with st.spinner("画像を解析中です...（枚数によって数十秒〜数分かかります）"):
            try:
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel(
                    'gemini-3.6-flash',
                    generation_config={"response_mime_type": "application/json"}
                )

                image_data_list = []
                file_name_mapping = []
                
                for idx, file in enumerate(uploaded_files):
                    img = PIL.Image.open(file)
                    image_data_list.append(img)
                    file_name_mapping.append(f"画像{idx+1}: {file.name}")
                
                mapping_text = "\n".join(file_name_mapping)

                prompt = f"""
                以下の売上日計表の画像（{len(uploaded_files)}枚）を読み取り、各画像について数値を照合し、JSONの配列形式で出力してください。
                
                【アップロードされたファイル名と画像の順序】
                {mapping_text}
                
                【基本照合項目（左側の「【売上集計表】レジ精算総計」と、右側の集計欄の数値を比較）】
                1.現金在高  2.過不足  3.商品券  4.キャッシュレス  5.入金額
                
                【追加の照合条件（ミスパンチ・返金）】
                ・売上日計表に記載されている「ミスパンチ」および「返金」の金額を確認してください。
                ・「ミスパンチ」の金額が0円でない（数値が存在する）場合：「返金ミスパンチ貼付用紙」の「ミスパンチ合計」と数値を照合し、一致するか確認してください。
                ・「返金」の金額が0円でない（数値が存在する）場合：「返金ミスパンチ貼付用紙」の「返金合計」と数値を照合し、一致するか確認してください。
                ・金額が0円または記載がない場合は、ステータスを「対象外(0円)」としてください。
                
                【重要な指示】
                ・同一内容（同じ店舗、同じ日付、同じ数値）の画像が複数含まれている場合は重複を排除し、1つのデータとしてまとめてください。
                ・「ファイル名」の値には、上記のファイル名リストを参照し、該当する正しいファイル名を記載してください（重複時はカンマ区切り）。
                
                【出力フォーマット】
                [
                  {{
                    "ファイル名": "...", 
                    "店舗名": "店舗名", 
                    "日付": "YYYY/MM/DD", 
                    "現金在高": "一致/不一致",
                    "過不足": "一致/不一致",
                    "商品券": "一致/不一致",
                    "キャッシュレス": "一致/不一致",
                    "入金額": "一致/不一致",
                    "ミスパンチ照合": "一致/不一致/対象外(0円)",
                    "返金照合": "一致/不一致/対象外(0円)",
                    "不一致の内容": "詳細（すべて一致・対象外の場合は 'なし'）"
                  }}
                ]
                """

                contents = [prompt] + image_data_list
                response = model.generate_content(contents)
                result_data = json.loads(response.text)

                st.success("✅ 解析および突合が完了しました！")

                df_results = pd.DataFrame(result_data)

                # --- 未提出店舗の抽出ロジック ---
                df_unsubmitted = pd.DataFrame()
                if master_df is not None and "店舗名" in master_df.columns:
                    extracted_stores = df_results["店舗名"].astype(str).str.replace("ジェーソン", "").str.strip().unique()
                    df_unsubmitted = master_df[~master_df["店舗名"].astype(str).str.strip().isin(extracted_stores)].copy()
                    df_unsubmitted["提出ステータス"] = "未提出"

                # --- 画面表示（タブ分け） ---
                tab1, tab2 = st.tabs(["📊 照合結果（提出分）", f"⚠️ 未提出店舗一覧（{len(df_unsubmitted)}店舗）" if not df_unsubmitted.empty else "⚠️ 未提出店舗一覧"])

                def highlight_mismatch(val):
                    color = '#ffcccc' if val == '不一致' else ''
                    return f'background-color: {color}'

                with tab1:
                    st.subheader("提出分・数値照合結果（ミスパンチ・返金照合含む）")
                    st.dataframe(df_results.style.map(highlight_mismatch), use_container_width=True)

                with tab2:
                    st.subheader("未提出店舗リスト")
                    if not df_unsubmitted.empty:
                        st.warning(f"全{len(master_df)}店舗中、**{len(df_unsubmitted)}店舗** の画像が未提出です。")
                        st.dataframe(df_unsubmitted, use_container_width=True)
                    else:
                        st.info("すべての対象店舗の画像が提出されています。")

                # --- Excelファイルの生成 ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_results.to_excel(writer, index=False, sheet_name="照合結果")
                    if not df_unsubmitted.empty:
                        df_unsubmitted.to_excel(writer, index=False, sheet_name="未提出店舗一覧")

                output.seek(0)
                wb = openpyxl.load_workbook(output)

                header_fill = PatternFill(start_color="34495E", end_color="34495E", fill_type="solid")
                header_font = Font(color="FFFFFF", bold=True)
                alignment_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
                alignment_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
                thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), 
                                     top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
                red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
                alt_fill = PatternFill(start_color="F9F9F9", end_color="F9F9F9", fill_type="solid")

                # シート1の装飾
                ws1 = wb["照合結果"]
                for cell in ws1[1]:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = alignment_center
                    cell.border = thin_border

                for row_idx, row in enumerate(ws1.iter_rows(min_row=2, max_row=ws1.max_row, min_col=1, max_col=ws1.max_column), start=2):
                    for col_idx, cell in enumerate(row):
                        cell.border = thin_border
                        if col_idx in [0, 1, 10]:  # ファイル名、店舗名、不一致の内容は左寄せ
                            cell.alignment = alignment_left
                        else:
                            cell.alignment = alignment_center
                        if row_idx % 2 == 0:
                            cell.fill = alt_fill
                        if cell.value == "不一致":
                            cell.fill = red_fill

                widths1 = {'A': 20, 'B': 25, 'C': 15, 'D': 12, 'E': 12, 'F': 12, 'G': 15, 'H': 12, 'I': 15, 'J': 15, 'K': 50}
                for col, width in widths1.items():
                    ws1.column_dimensions[col].width = width

                # シート2の装飾
                if "未提出店舗一覧" in wb.sheetnames:
                    ws2 = wb["未提出店舗一覧"]
                    for cell in ws2[1]:
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = alignment_center
                        cell.border = thin_border

                    for row_idx, row in enumerate(ws2.iter_rows(min_row=2, max_row=ws2.max_row, min_col=1, max_col=ws2.max_column), start=2):
                        for col_idx, cell in enumerate(row):
                            cell.border = thin_border
                            cell.alignment = alignment_left
                            if row_idx % 2 == 0:
                                cell.fill = alt_fill

                    widths2 = {'A': 15, 'B': 25, 'C': 15}
                    for col, width in widths2.items():
                        ws2.column_dimensions[col].width = width

                final_output = io.BytesIO()
                wb.save(final_output)
                excel_data = final_output.getvalue()

                st.download_button(
                    label="📥 照合結果＆未提出リスト（Excel）をダウンロード",
                    data=excel_data,
                    file_name="売上照合および未提出店舗報告書.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"システムエラーが発生しました: {e}")
