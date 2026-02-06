import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import datetime
import xlsxwriter
import openpyxl

# ================= 0. 系統設定 =================
st.set_page_config(page_title="科普列車統計系統 V10.0", page_icon="🎯", layout="wide")

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

def process_file_custom(filename, content, password, target_cols):
    """
    V10.0 核心：依據使用者輸入的欄位名稱進行抓取
    target_cols = {'city': '使用者輸入的縣市欄位', 'school': '使用者輸入的學校欄位', ...}
    """
    wb = open_excel_safe(content, password)
    if wb is None:
        return None, {"filename": filename, "status": "Fail", "msg": "無法開啟(密碼錯誤)"}

    ws = wb.active
    
    # 1. 找表頭 (還是需要自動判斷哪一行是標題)
    # 我們假設表頭會包含使用者輸入的「學校」或「縣市」欄位名稱
    header_row_idx = 0
    found_header = False
    
    # 掃描前 10 行
    for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True)):
        row_str = [str(c).strip() for c in row if c is not None]
        # 如果這一行包含了使用者指定的欄位名稱，就認定是表頭
        if target_cols['school'] in row_str or target_cols['city'] in row_str:
            header_row_idx = r_idx
            found_header = True
            break
            
    if not found_header:
        return None, {"filename": filename, "status": "Fail", "msg": f"找不到表頭 (請確認Excel內是否有 '{target_cols['school']}' 或 '{target_cols['city']}' 欄位)"}

    # 2. 讀取資料
    data = list(ws.values)
    raw_header = data[header_row_idx]
    rows = data[header_row_idx+1:]
    
    df = pd.DataFrame(rows, columns=raw_header)
    
    # 清理欄位名稱 (去空白)
    df.columns = [str(c).strip() for c in df.columns]
    
    # 3. 檢查使用者指定的欄位是否存在
    missing_cols = []
    clean_data = {}
    
    # 定義我們要抓的三大維度
    mapping = {
        'city': '縣市',
        'school': '學校',
        'role': '職稱'
    }

    for user_key, output_name in mapping.items():
        user_input_col_name = target_cols.get(user_key)
        
        if user_input_col_name in df.columns:
            clean_data[output_name] = df[user_input_col_name]
        else:
            clean_data[output_name] = "未找到欄位"
            missing_cols.append(user_input_col_name)

    # 建立統計用 DataFrame
    df_stat = pd.DataFrame(clean_data)
    df_stat.fillna("未知", inplace=True)
    
    # 記錄訊息
    msg = "OK"
    if missing_cols:
        msg = f"部分欄位未找到: {', '.join(missing_cols)}"
    
    return df_stat, {"filename": filename, "status": "Success" if not missing_cols else "Warning", "msg": msg}

# ================= 2. 批量執行 =================

def run_analysis(files, pwd, user_cols):
    log_list = []
    combined_df = pd.DataFrame() 
    
    progress_bar = st.progress(0)
    
    for i, f in enumerate(files):
        try:
            df_stat, meta = process_file_custom(f.name, f.read(), pwd, user_cols)
            log_list.append(meta)
            
            if df_stat is not None:
                df_stat['來源檔案'] = f.name
                combined_df = pd.concat([combined_df, df_stat], ignore_index=True)
        except Exception as e:
            st.error(f"檔案 {f.name} 處理失敗: {e}")
            
        progress_bar.progress((i + 1) / len(files))
        
    return log_list, combined_df

# ================= 3. 主介面 =================

st.title("🎯 科普列車 - 自定義欄位分析 V10.0")
st.markdown("### 請告訴系統，您的 Excel 欄位叫什麼名字？")

# --- 左側設定區 ---
with st.sidebar:
    st.header("⚙️ 欄位設定")
    st.info("請輸入
