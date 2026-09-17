import streamlit as st
import google.generativeai as genai
import PIL.Image
import pandas as pd
import io
import json
import os
import datetime
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

# 1. 店舗マスタの読み込み
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

# 2. 調査対象日の設定
st.subheader("2. 調査対象日の設定")
col1, col2 = st.columns([1, 2])
with col1:
    ignore_date_check = st.checkbox("調査対象日を指定しない（すべての画像を取り込む）", value=False)
with col2:
    target_date = st.date_input("調査対象日を選択してください", datetime.date.today(), disabled=ignore_date_check)

target_date_str = "指定なし（全日付を照合対象とする）" if ignore_date_check else target_date.strftime("%Y/%m/%d")

# 3. 画像アップロード
st.subheader("3. 日計表画像のアップロード")
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
                all_uploaded_filenames = [f.name for f in uploaded_files]
                
                for idx, file in enumerate(uploaded_files):
                    img = PIL.Image.open(file)
                    image_data_list.append(img)
                    file_name_mapping.append(f"画像{idx+1}: {file.name}")
                
                mapping_text = "\n".join(file_name_mapping)

                prompt = f"""
                以下の売上日計表の画像（{len(uploaded_files)}枚）を読み取り、各画像について数値を照合し、JSONオブジェクト形式で出力してください。
                
                【アップロードされたファイル名と画像の順序】
                {mapping_text}
                
                【重要：判別不能・不明画像の無視（スキップ）ルール】
                ・画像がブレている、極度に不鮮明、見切れている、または売上日計表ではないなどの理由で**数値や店舗名が判別・判断できない画像**は、照合対象から完全に除外（無視）してください。
                ・除外した画像のファイル名は `unreadable_files` 配列に記載してください。
                
                【重要：調査対象日フィルター規則】
                ・指定の調査対象日：{target_date_str}
                ・「{target_date_str}」が「指定なし」でない場合、画像内の売上日計表の日付（年月日）が対象日と一致するか判定してください。
                ・日付が一致しない場合、各照合項目を「対象外(日付不一致)」とし、不一致の内容に「指定された調査対象日（{target_date_str}）と異なります」と記録してください。
                
                【基本照合項目（左側の「【売上集計表】レジ精算総計」と、右側の集計欄の数値を比較）】
                1.現金在高  2.過不足  3.商品券  4.キャッシュレス  5.入金額
                
                【追加の照合条件（ミスパンチ・返金）】
                ・売上日計表に記載されている「ミスパンチ」および「返金」の金額を確認してください。
                ・「ミスパンチ」の金額が0円でない場合：「返金ミスパンチ貼付用紙」の「ミスパンチ合計」と数値を照合してください。
                ・「返金」の金額が0円でない場合：「返金ミスパンチ貼付用紙」の「返金合計」と数値を照合してください。
                ・金額が0円または記載がない場合は、ステータスを「対象外(0円)」としてください。
                
                【出力フォーマット】
                {{
                  "results": [
                    {{
                      "ファイル名": "...", 
                      "店舗名": "店舗名", 
                      "日付": "YYYY/MM/DD", 
                      "現金在高": "一致/不一致/対象外(日付不一致)",
                      "過不足": "一致/不一致/対象外(日付不一致)",
                      "商品券": "一致/不一致/対象外(日付不一致)",
                      "キャッシュレス": "一致/不一致/対象外(日付不一致)",
                      "入金額": "一致/不一致/対象外(日付不一致)",
                      "ミスパンチ照合": "一致/不一致/対象外(0円)/対象外(日付不一致)",
                      "返金照合": "一致/不一致/対象外(0円)/対象外(日付不一致)",
                      "不一致の内容": "詳細（すべて一致・対象外の場合は 'なし'）"
                    }}
                  ],
                  "unreadable_files": ["判別不能なファイル名1.jpg", "..."]
                }}
                """

                contents = [prompt] + image_data_list
                response = model.generate_content(contents)
                raw_json = json.loads(response.text)

                # 結果データと判別不能ファイルリストの分離
                if isinstance(raw_json, dict):
                    result_data = raw_json.get("results", [])
                    unreadable_files = raw_json.get("unreadable_files", [])
                else:
                    result_data = raw_json
                    unreadable_files = []

                st.success("✅ 解析および突合が完了しました！")

                df_results = pd.DataFrame(result_data)
                df_unreadable = pd.DataFrame({"判別不能ファイル名": unreadable_files, "理由": "画像不鮮明・判別不能につき無視"}) if unreadable_files else pd.DataFrame()

                # --- 未提出店舗の抽出ロジック（対象日一致データのみで集計） ---
                df_unsubmitted = pd.DataFrame()
                if master_df is not None and "店舗名" in master_df.columns:
                    if not df_results.empty and "現金在高" in df_results.columns:
                        valid_results = df_results[~df_results["現金在高"].astype(str).str.contains("日付不一致")]
                        extracted_stores = valid_results["店舗名"].astype(str).str.replace("ジェーソン", "").str.strip().unique()
                    else:
                        extracted_stores = []

                    df_unsubmitted = master_df[~master_df["店舗名"].astype(str).str.strip().isin(extracted_stores)].copy()
                    df_unsubmitted["提出ステータス"] = "未提出"

                # --- 画面表示（タブ分け） ---
                tab1, tab2, tab3 = st.tabs([
                    "📊 照合結果（提出分）", 
                    f"⚠️ 未提出店舗一覧（{len(df_unsubmitted)}店舗）" if not df_unsubmitted.empty else "⚠️ 未提出店舗一覧",
                    f"🚫 判別不能・無視ファイル（{len(unreadable_files)}件）" if unreadable_files else "🚫 判別不能・無視ファイル"
                ])

                def highlight_mismatch(val):
                    if val == '不一致':
                        return 'background-color: #ffcccc'
                    elif '日付不一致' in str(val):
                        return 'background-color: #eeeeee; color: #888888'
                    return ''

                with tab1:
                    st.subheader(f"提出分・数値照合結果（対象日: {target_date_str}）")
                    if not df_results.empty:
                        st.dataframe(df_results.style.map(highlight_mismatch), use_container_width=True)
                    else:
                        st.info("有効な照合対象データがありませんでした。")

                with tab2:
                    st.subheader("未提出店舗リスト")
                    if not df_unsubmitted.empty:
                        st.warning(f"全{len(master_df)}店舗中、対象日（{target_date_str}）の画像が **{len(df_unsubmitted)}店舗** 未提出です。")
                        st.dataframe(df_unsubmitted, use_container_width=True)
                    else:
                        st.info("すべての対象店舗の画像が提出されています。")

                with tab3:
                    st.subheader("判別不能・無視されたファイル")
                    if unreadable_files:
                        st.error(f"以下の **{len(unreadable_files)}件** の画像は画質不良または不整合のため判別できず、照合から除外されました。")
                        st.dataframe(df_unreadable, use_container_width=True)
                    else:
                        st.success("判別不能・無視された画像はありませんでした。すべての画像が正常に処理されました。")

                # --- Excelファイルの生成 ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_results.to_excel(writer, index=False, sheet_name="照合結果")
                    if not df_unsubmitted.empty:
                        df_unsubmitted.to_excel(writer, index=False, sheet_name="未提出店舗一覧")
                    if not df_unreadable.empty:
                        df_unreadable.to_excel(writer, index=False, sheet_name="判別不能ファイル一覧")

                output.seek(0)
                wb = openpyxl.load_workbook(output)

                header_fill = PatternFill(start_color="34495E", end_color="34495E", fill_type="solid")
                header_font = Font(color="FFFFFF", bold=True)
                alignment_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
                alignment_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
                thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), 
                                     top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
                red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
                gray_fill = PatternFill(start_color="EAEAEA", end_color="EAEAEA", fill_type="solid")
                alt_fill = PatternFill(start_color="F9F9F9", end_color="F9F9F9", fill_type="solid")

                # シートの装飾共通処理
                for sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    for cell in ws[1]:
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = alignment_center
                        cell.border = thin_border

                    for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column), start=2):
                        for col_idx, cell in enumerate(row):
                            cell.border = thin_border
                            if sheet_name == "照合結果" and col_idx not in [0, 1, 10]:
                                cell.alignment = alignment_center
                            else:
                                cell.alignment = alignment_left

                            if row_idx % 2 == 0:
                                cell.fill = alt_fill

                            if cell.value == "不一致":
                                cell.fill = red_fill
                            elif cell.value == "対象外(日付不一致)":
                                cell.fill = gray_fill

                final_output = io.BytesIO()
                wb.save(final_output)
                excel_data = final_output.getvalue()

                st.download_button(
                    label="📥 照合結果・未提出・不明ファイルリスト（Excel）をダウンロード",
                    data=excel_data,
                    file_name="売上照合および未提出店舗報告書.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"システムエラーが発生しました: {e}")
