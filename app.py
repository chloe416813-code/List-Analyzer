import streamlit as st
import pandas as pd
import io
import zipfile
import xlsxwriter
import openpyxl

# ================= 0. 系統設定 =================
st.set_page_config(page_title="萬用 Excel 統計系統 V11.0", page_icon="🧩", layout="wide")

# 初始化 Session State
if 'analysis_done' not in st.session_state:
    st.session_state.analysis_done = False
if 'stats_excel' not in st.session_state:
    st.session_state.stats_excel = None
if 'big_df' not in st.session_state:
    st.session_state.big_df = pd.DataFrame()
if 'log_report' not in st.session_state:
    st.session_state.log_report = []

# ================= 1. 核心邏輯區 =================

def open_excel_safe(file_content, password):
    """ 安全開啟 Excel (支援加密) """
    file_stream = io.BytesIO(file_content)
    try:
        return openpyxl.load_workbook(file_stream, data_only=True)
    except:
        file_stream.seek(0)
    
    if password:
        try:
            import msoffcrypto
            decrypted = io.BytesIO()
            office_file = msoffcrypto.OfficeFile(file_stream)
            office_file.load_key(password=password)
            office_file.decrypt(decrypted)
            decrypted.seek(0)
            return openpyxl.load_workbook(decrypted, data_only=True)
        except:
            return None
    return None

def process_file_dynamic(filename, content, password, target_cols_list):
    """
    V11.0 核心：依據使用者在表格中設定的欄位清單進行抓取
    """
    wb = open_excel_safe(content, password)
    if wb is None:
        return None, {"filename": filename, "status": "Fail", "msg": "無法開啟(密碼錯誤)"}

    ws = wb.active
    
    # 1. 找表頭
    # 邏輯：掃描前 10 行，只要該行包含了使用者指定欄位的「任意一個」，就認定是表頭
    header_row_idx = 0
    found_header = False
    
    # 為了提高命中率，我們把使用者輸入的欄位都轉字串並去空白
    target_set = set([str(c).strip() for c in target_cols_list])
    
    for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True)):
        # 讀取這一行的內容，轉成乾淨的清單
        row_values = [str(c).strip() for c in row if c is not None]
        
        # 檢查是否有交集 (這一行是否包含使用者要的欄位)
        # 只要命中 1 個，我們就假設這是表頭行
        if set(row_values).intersection(target_set):
            header_row_idx = r_idx
            found_header = True
            break
            
    if not found_header:
        return None, {"filename": filename, "status": "Fail", "msg": f"找不到符合設定的表頭，請確認 Excel 欄位名稱。"}

    # 2. 讀取資料
    data = list(ws.values)
    raw_header = data[header_row_idx]
    rows = data[header_row_idx+1:]
    
    df = pd.DataFrame(rows, columns=raw_header)
    
    # 清理 DataFrame 欄位名稱
    df.columns = [str(c).strip() for c in df.columns]
    
    # 3. 萃取使用者指定的欄位
    clean_data = {}
    missing_cols = []

    for col_name in target_cols_list:
        clean_col = str(col_name).strip()
        
        # 嘗試在 Excel 中找這個欄位 (完全匹配)
        if clean_col in df.columns:
             clean_data[clean_col] = df[clean_col]
        else:
             # 嘗試模糊匹配 (例如使用者設「學校」，可以抓到「就讀學校」)
             found_fuzzy = False
             for excel_col in df.columns:
                 if clean_col in excel_col:
                     clean_data[clean_col] = df[excel_col] # 統一使用使用者設定的名稱當 Key
                     found_fuzzy = True
                     break
             
             if not found_fuzzy:
                 clean_data[clean_col] = "未找到"
                 missing_cols.append(clean_col)

    df_stat = pd.DataFrame(clean_data)
    df_stat.fillna("", inplace=True)
    
    msg = "OK"
    status = "Success"
    if missing_cols:
        msg = f"部分未找到: {', '.join(missing_cols)}"
        status = "Warning"
    
    return df_stat, {"filename": filename, "status": status, "msg": msg}

# ================= 2. 批量執行 =================

def run_analysis(files, pwd, target_cols_list):
    log_list = []
    combined_df = pd.DataFrame() 
    
    progress_bar = st.progress(0)
    
    for i, f in enumerate(files):
        try:
            df_stat, meta = process_file_dynamic(f.name, f.read(), pwd, target_cols_list)
            log_list.append(meta)
            
            if df_stat is not None:
                df_stat['來源檔案'] = f.name
                combined_df = pd.concat([combined_df, df_stat], ignore_index=True)
        except Exception as e:
            st.error(f"檔案 {f.name} 處理失敗: {e}")
            
        progress_bar.progress((i + 1) / len(files))
        
    return log_list, combined_df

# ================= 3. 主介面 =================

st.title("🧩 萬用 Excel 統計系統 V11.0")
st.markdown("### 完全自定義：您想抓什麼欄位，自己決定！")

# --- 左側設定區 ---
with st.sidebar:
    st.header("1. 設定分析欄位")
    st.info("請在下方表格新增您想要分析的 Excel 欄位名稱。")
    st.caption("例如：學校、縣市、職稱、金額...")
    
    # 初始化一個空的 DataFrame 給使用者輸入
    # 這就是讓使用者「自行增減」的關鍵元件
    default_data = pd.DataFrame([{"欄位名稱": ""}]) # 預設留空，或給一個空行
    
    edited_df = st.data_editor(
        pd.DataFrame(columns=["欄位名稱"]), # 預設完全空白
        num_rows="dynamic", # 允許使用者按 + 新增行
        key="col_editor",
        use_container_width=True
    )
    
    # 轉換使用者輸入為清單
    target_cols = [row["欄位名稱"] for i, row in edited_df.iterrows() if row["欄位名稱"]]
    
    st.divider()
    pwd_input = st.text_input("檔案密碼 (若無則留空)", type="password")

# --- 右側上傳區 ---
st.header("2. 上傳檔案與執行")
files_input = st.file_uploader("請上傳 Excel (支援多檔合併)", type=['xlsx'], accept_multiple_files=True)

if st.button("🚀 開始分析", type="primary"):
    if not files_input:
        st.warning("請先上傳檔案！")
    elif not target_cols:
        st.warning("請在左側側邊欄【新增】至少一個要分析的欄位名稱！")
    else:
        st.session_state.big_df = pd.DataFrame()
        
        with st.spinner(f"正在抓取欄位: {target_cols} ..."):
            log_list, big_df = run_analysis(files_input, pwd_input, target_cols)
            
            st.session_state.big_df = big_df
            st.session_state.log_report = log_list
            st.session_state.analysis_done = True

# ================= 4. 結果顯示與報表設定 =================

if st.session_state.analysis_done:
    st.divider()
    
    # 顯示處理日誌
    with st.expander("📄 處理狀態報告", expanded=False):
        st.dataframe(pd.DataFrame(st.session_state.log_report), use_container_width=True)

    if not st.session_state.big_df.empty:
        df = st.session_state.big_df
        
        st.subheader("3. 分析結果預覽")
        st.dataframe(df.head(), use_container_width=True)
        st.caption(f"共擷取 {len(df)} 筆資料，來自 {df['來源檔案'].nunique()} 個檔案。")
        
        st.divider()
        st.subheader("4. 產生統計報表")
        st.info("請選擇如何統計這些資料，系統將為您產生 Excel 樞紐分析表。")
        
        col1, col2 = st.columns(2)
        with col1:
            # 讓使用者選擇「列 (Index)」(例如：縣市、學校)
            pivot_index = st.multiselect(
                "選擇【分類】欄位 (Row)", 
                options=target_cols,
                default=target_cols[:2] if len(target_cols) >= 2 else target_cols
            )
        with col2:
            # 讓使用者選擇「欄 (Column)」(例如：職稱)
            pivot_columns = st.multiselect(
                "選擇【比較】欄位 (Column) (可選)", 
                options=[c for c in target_cols if c not in pivot_index]
            )
            
        if st.button("📊 產生並下載 Excel 報表"):
            try:
                output_io = io.BytesIO()
                with pd.ExcelWriter(output_io, engine='xlsxwriter') as writer:
                    
                    # 1. 總名單
                    df.to_excel(writer, sheet_name='合併總表', index=False)
                    
                    # 2. 自動統計 (計數)
                    # 針對使用者設定的每一個欄位，都做一個簡單的計數表
                    for col in target_cols:
                        if col in df.columns:
                            counts = df[col].value_counts().to_frame(name="數量")
                            counts.to_excel(writer, sheet_name=f'{col}統計')
                    
                    # 3. 使用者自訂的樞紐分析
                    if pivot_index:
                        # 這是 Pivot Table 的核心
                        pivot = df.pivot_table(
                            index=pivot_index, 
                            columns=pivot_columns[0] if pivot_columns else None, 
                            aggfunc='size', 
                            fill_value=0
                        )
                        
                        # 如果有比較欄位，加總計
                        if pivot_columns:
                            pivot['總計'] = pivot.sum(axis=1)
                            
                        pivot.to_excel(writer, sheet_name='自訂交叉分析')
                        
                        # 調整欄寬
                        writer.sheets['自訂交叉分析'].set_column(0, len(pivot_index), 20)
                
                output_io.seek(0)
                
                st.download_button(
                    label="📥 下載 Excel 完整報表",
                    data=output_io,
                    file_name="自訂統計報表.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                
            except Exception as e:
                st.error(f"報表產生錯誤 (可能是選取的欄位資料有誤): {e}")

    else:
        st.warning("沒有抓取到任何資料，請檢查欄位名稱設定。")
