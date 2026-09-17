import streamlit as st
import google.generativeai as genai
import PIL.Image
import pandas as pd
import io
import json
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
st.markdown("店舗から送られてきた売上日計表の画像をアップロードするだけで、数値を自動照合し、結果をExcelで出力します。")

# SecretsからAPIキーを取得
api_key = st.secrets.get("GEMINI_API_KEY")

uploaded_files = st.file_uploader(
    "売上日計表の画像をアップロードしてください（複数選択可・ドラッグ＆ドロップ対応）", 
    type=['png', 'jpg', 'jpeg'], 
    accept_multiple_files=True
)

if st.button("照合を開始する", type="primary"):
    if not api_key:
        st.error("システムエラー: APIキーが設定されていません。StreamlitのSecretsを設定してください。")
    elif not uploaded_files:
        st.warning("画像を1枚以上アップロードしてください。")
    else:
        with st.spinner("画像を解析中です...（枚数によって数十秒〜数分かかります）"):
            try:
                genai.configure(api_key=api_key)
                # エラー指示に従いモデル名を gemini-3.6-flash に変更
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
                以下の売上日計表の画像（{len(uploaded_files)}枚）を読み取り、各画像について5項目を照合し、JSONの配列形式で出力してください。
                
                【アップロードされたファイル名と画像の順序】
                {mapping_text}
                
                【照合項目（左側の「【売上集計表】レジ精算総計」と、右側の集計欄の数値を比較）】
                1.現金在高  2.過不足  3.商品券  4.キャッシュレス  5.入金額
                
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
                    "不一致の内容": "詳細（すべて一致の場合は 'なし'）"
                  }}
                ]
                """

                contents = [prompt] + image_data_list
                response = model.generate_content(contents)
                result_data = json.loads(response.text)

                st.success("✅ 解析が完了しました！")

                df = pd.DataFrame(result_data)
                
                def highlight_mismatch(val):
                    color = '#ffcccc' if val == '不一致' else ''
                    return f'background-color: {color}'
                
                st.subheader("照合結果")
                st.dataframe(df.style.map(highlight_mismatch), use_container_width=True)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name="照合結果")
                
                output.seek(0)
                wb = openpyxl.load_workbook(output)
                ws = wb["照合結果"]

                header_fill = PatternFill(start_color="34495E", end_color="34495E", fill_type="solid")
                header_font = Font(color="FFFFFF", bold=True)
                alignment_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
                alignment_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
                thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), 
                                     top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
                red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
                alt_fill = PatternFill(start_color="F9F9F9", end_color="F9F9F9", fill_type="solid")

                for cell in ws[1]:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = alignment_center
                    cell.border = thin_border

                for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column), start=2):
                    for col_idx, cell in enumerate(row):
                        cell.border = thin_border
                        if col_idx in [0, 1, 8]:
                            cell.alignment = alignment_left
                        else:
                            cell.alignment = alignment_center
                        if row_idx % 2 == 0:
                            cell.fill = alt_fill
                        if cell.value == "不一致":
                            cell.fill = red_fill

                widths = {'A': 20, 'B': 25, 'C': 15, 'D': 12, 'E': 12, 'F': 12, 'G': 15, 'H': 12, 'I': 50}
                for col, width in widths.items():
                    ws.column_dimensions[col].width = width

                final_output = io.BytesIO()
                wb.save(final_output)
                excel_data = final_output.getvalue()

                st.download_button(
                    label="📥 照合結果（Excelファイル）をダウンロード",
                    data=excel_data,
                    file_name="売上照合結果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"システムエラーが発生しました: {e}")
