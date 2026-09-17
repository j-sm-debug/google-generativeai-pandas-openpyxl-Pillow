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

# --- セッション状態（データ蓄積用メモリ・キー）の初期化 ---
if "accumulated_results" not in st.session_state:
    st.session_state["accumulated_results"] = pd.DataFrame()
if "accumulated_unreadable" not in st.session_state:
    st.session_state["accumulated_unreadable"] = pd.DataFrame()
if "file_uploader_key" not in st.session_state:
    st.session_state["file_uploader_key"] = 0

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

# --- 店舗名自動補正関数 ---
def normalize_and_fix_store_names(df_results, master_df):
    if df_results.empty or master_df is None or "店舗コード" not in master_df.columns or "店舗名" not in master_df.columns:
        return df_results
    
    code_to_name = {}
    master_store_names = []
    
    for _, row in master_df.iterrows():
        raw_code = str(row["店舗コード"]).strip()
        raw_name = str(row["店舗名"]).strip().replace("ジェーソン", "")
        master_store_names.append(raw_name)
        
        code_to_name[raw_code] = raw_name
        if raw_code.isdigit():
            code_to_name[raw_code.zfill(3)] = raw_name
            code_to_name[str(int(raw_code))] = raw_name

    fixed_stores = []
    
    for idx, row in df_results.iterrows():
        fname = str(row.get("ファイル名", ""))
        ocr_store = str(row.get("店舗名", "")).strip()
        
        base_name = os.path.splitext(os.path.basename(fname))[0].strip()
        code_prefix = base_name.split('-')[0].split('_')[0].split(' ')[0].strip()
        
        matched_name = None
        
        if code_prefix in code_to_name:
            matched_name = code_to_name[code_prefix]
            
        if not matched_name:
            for code, name in code_to_name.items():
                if len(code) >= 2 and (base_name.startswith(code) or f"-{code}" in base_name or f"_{code}" in base_name):
                    matched_name = name
                    break
                    
        if not matched_name:
            clean_ocr = ocr_store.replace("ジェーソン", "").strip()
            if clean_ocr in master_store_names:
                matched_name = clean_ocr
        
        if matched_name:
            fixed_stores.append(f"ジェーソン{matched_name}")
        else:
            fixed_stores.append(ocr_store)
            
    df_results["店舗名"] = fixed_stores
    return df_results

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
    st.warning("⚠️ 店舗マスタ（店舗.xlsx）が読み込まれていません。")

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
    "売上日計表の画像をアップロードしてください（複数選択可・順次追加可能）", 
    type=['png', 'jpg', 'jpeg'], 
    accept_multiple_files=True,
    key=f"uploader_{st.session_state['file_uploader_key']}"
)

# ボタンの3列レイアウト配置
btn_col1, btn_col2, btn_col3 = st.columns([1.5, 1.5, 1.5])

with btn_col1:
    start_btn = st.button("🚀 照合を追加実行する", type="primary", use_container_width=True)

with btn_col2:
    clear_file_btn = st.button("📁 アップロード選択欄をクリア（結果は残す）", use_container_width=True)

with btn_col3:
    clear_all_btn = st.button("🗑️ 全調査完了（全結果データをクリアする）", use_container_width=True)

# --- アップロード選択欄のみクリアする処理 ---
if clear_file_btn:
    st.session_state["file_uploader_key"] += 1
    st.rerun()

# --- 全結果クリアの処理 ---
if clear_all_btn:
    st.session_state["accumulated_results"] = pd.DataFrame()
    st.session_state["accumulated_unreadable"] = pd.DataFrame()
    st.session_state["file_uploader_key"] += 1
    st.success("🧹 照合結果データをすべてリセットしました。")
    st.rerun()

# --- 照合処理実行 ---
if start_btn:
    if not api_key:
        st.error("システムエラー: APIキーが設定されていません。StreamlitのSecretsを設定してください。")
    elif not uploaded_files:
        st.warning("追加する画像を1枚以上アップロードしてください。")
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

                master_info_str = ""
                if master_df is not None and "店舗コード" in master_df.columns and "店舗名" in master_df.columns:
                    master_info_str = "【店舗マスタ（ファイル名の数字/記号は店舗コードに対応しています）】\n"
                    for _, r in master_df.iterrows():
                        master_info_str += f"- コード `{r['店舗コード']}`: {r['店舗名']}\n"

                prompt = f"""
                以下の売上日計表の画像（{len(uploaded_files)}枚）を読み取り、各画像について数値を照合し、JSONオブジェクト形式で出力してください。
                
                【アップロードされたファイル名と画像の順序】
                {mapping_text}
                
                {master_info_str}
                
                【重要：店舗名の判定規則】
                ・ファイル名の先頭部分（例: `B75.jpg` -> `B75`, `144-1.jpg` -> `144`）は店舗マスタの店舗コードに対応しています。
                ・店舗名は必ず店舗マスタに記載された正確な店舗名（例: `ジェーソン和光店`）を判定してください。
                
                【重要：判別不能・不明画像の無視（スキップ）ルール】
                ・画像がブレている、極度に不鮮明、見切れている、または売上日計表ではないなどの理由で数値や店舗名が判別できない画像は照合対象から除外してください。
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

                if isinstance(raw_json, dict):
                    result_data = raw_json.get("results", [])
                    unreadable_files = raw_json.get("unreadable_files", [])
                else:
                    result_data = raw_json
                    unreadable_files = []

                new_df_results = pd.DataFrame(result_data)
                
                # 店舗名の自動補正
                if not new_df_results.empty:
                    new_df_results = normalize_and_fix_store_names(new_df_results, master_df)
                    for col in new_df_results.columns:
                        new_df_results[col] = new_df_results[col].astype(str)

                new_df_unreadable = pd.DataFrame({"判別不能ファイル名": unreadable_files, "理由": "画像不鮮明・判別不能につき無視"}) if unreadable_files else pd.DataFrame()
                if not new_df_unreadable.empty:
                    for col in new_df_unreadable.columns:
                        new_df_unreadable[col] = new_df_unreadable[col].astype(str)

                # --- 既存データへの追記と重複排除 ---
                if not new_df_results.empty:
                    combined_res = pd.concat([st.session_state["accumulated_results"], new_df_results], ignore_index=True)
                    st.session_state["accumulated_results"] = combined_res.drop_duplicates(subset=["ファイル名"], keep="first")

                if not new_df_unreadable.empty:
                    combined_unread = pd.concat([st.session_state["accumulated_unreadable"], new_df_unreadable], ignore_index=True)
                    st.session_state["accumulated_unreadable"] = combined_unread.drop_duplicates(subset=["判別不能ファイル名"], keep="first")

                st.success("✅ 解析結果を追加・更新しました！")

            except Exception as e:
                st.error(f"システムエラーが発生しました: {e}")

# --- 画面表示部（累積データを表示） ---
df_results = st.session_state["accumulated_results"]
df_unreadable = st.session_state["accumulated_unreadable"]

if not df_results.empty or not df_unreadable.empty:
    # 未提出店舗の抽出
    df_unsubmitted = pd.DataFrame()
    if master_df is not None and "店舗名" in master_df.columns:
        if not df_results.empty and ("現金を高" in df_results.columns or "現金在高" in df_results.columns):
            col_name = "現金在高" if "現金在高" in df_results.columns else "現金を高"
            valid_results = df_results[~df_results[col_name].astype(str).str.contains("日付不一致")]
            extracted_stores = valid_results["店舗名"].astype(str).str.replace("ジェーソン", "").str.strip().unique()
        else:
            extracted_stores = []

        df_unsubmitted = master_df[~master_df["店舗名"].astype(str).str.strip().isin(extracted_stores)].copy()
        df_unsubmitted["提出ステータス"] = "未提出"
        for col in df_unsubmitted.columns:
            df_unsubmitted[col] = df_unsubmitted[col].astype(str)

    # 画面表示（タブ分け）
    tab1, tab2, tab3 = st.tabs([
        f"📊 照合結果（累計: {len(df_results)}件）", 
        f"⚠️ 未提出店舗一覧（{len(df_unsubmitted)}店舗）" if not df_unsubmitted.empty else "⚠️ 未提出店舗一覧",
        f"🚫 判別不能・無視ファイル（{len(df_unreadable)}件）" if not df_unreadable.empty else "🚫 判別不能・無視ファイル"
    ])

    def highlight_mismatch(val):
        if val == '不一致':
            return 'background-color: #ffcccc'
        elif '日付不一致' in str(val):
            return 'background-color: #eeeeee; color: #888888'
        return ''

    with tab1:
        st.subheader(f"提出分・累計数値照合結果（対象日: {target_date_str}）")
        if not df_results.empty:
            st.dataframe(df_results.style.map(highlight_mismatch), use_container_width=True)

    with tab2:
        st.subheader("未提出店舗リスト")
        if not df_unsubmitted.empty:
            st.warning(f"全{len(master_df)}店舗中、対象日（{target_date_str}）の画像が **{len(df_unsubmitted)}店舗** 未提出です。")
            st.dataframe(df_unsubmitted, use_container_width=True)
        else:
            st.info("すべての対象店舗の画像が提出されています。")

    with tab3:
        st.subheader("判別不能・無視されたファイル")
        if not df_unreadable.empty:
            st.error(f"以下の **{len(df_unreadable)}件** の画像は画質不良または不整合のため判別できず除外されています。")
            st.dataframe(df_unreadable, use_container_width=True)
        else:
            st.success("判別不能・無視された画像はありません。")

    # Excelファイル生成
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if not df_results.empty:
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
        label="📥 累計照合結果・未提出・不明ファイルリスト（Excel）をダウンロード",
        data=excel_data,
        file_name="売上照合および未提出店舗報告書_累計.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
