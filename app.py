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
    st.info("請輸入您 Excel 檔案中【第一列表頭】的準確名稱。")
    
    col_city_input = st.text_input("1. [縣市] 欄位名稱", value="居住縣市")
    col_school_input = st.text_input("2. [學校] 欄位名稱", value="就讀學校")
    col_role_input = st.text_input("3. [職稱/身分] 欄位名稱", value="身分")
    
    st.divider()
    pwd_input = st.text_input("檔案密碼 (若無則留空)", type="password")

# --- 右側上傳區 ---
files_input = st.file_uploader("請上傳 Excel (支援多檔合併)", type=['xlsx'], accept_multiple_files=True)

if st.button("🚀 依指定欄位開始分析", type="primary"):
    if not files_input:
        st.warning("請先上傳檔案！")
    else:
        # 包裝使用者設定
        user_cols_config = {
            'city': col_city_input.strip(),
            'school': col_school_input.strip(),
            'role': col_role_input.strip()
        }
        
        st.session_state.big_df = pd.DataFrame()
        
        with st.spinner("正在依照您的指令抓取資料..."):
            log_list, big_df = run_analysis(files_input, pwd_input, user_cols_config)
            
            if not big_df.empty:
                try:
                    stats_io = io.BytesIO()
                    # 樞紐分析
                    # 確保沒有全空的狀況
                    if '縣市' in big_df.columns and '學校' in big_df.columns:
                        pivot = big_df.pivot_table(index=['縣市', '學校'], columns='職稱', aggfunc='size', fill_value=0)
                        pivot['該校總計'] = pivot.sum(axis=1)
                        
                        with pd.ExcelWriter(stats_io, engine='xlsxwriter') as writer:
                            pivot.to_excel(writer, sheet_name='各校統計')
                            big_df['縣市'].value_counts().to_frame(name="人數").to_excel(writer, sheet_name='縣市統計')
                            big_df.to_excel(writer, sheet_name='總名單明細', index=False)
                            
                            writer.sheets['各校統計'].set_column(0, 1, 20)
                            
                        st.session_state.stats_excel = stats_io.getvalue()
                        st.session_state.big_df = big_df
                    else:
                        st.error("無法產生報表，因為找不到必要的統計欄位。")
                    
                except Exception as e:
                    st.error(f"報表產生失敗: {e}")
            
            st.session_state.log_report = log_list
            st.session_state.analysis_done = True

# ================= 4. 結果顯示區 =================

if st.session_state.analysis_done:
    st.divider()
    
    # 顯示處理日誌 (讓使用者知道有沒有抓錯)
    with st.expander("📄 查看處理狀態 (欄位是否抓取成功?)", expanded=True):
        status_df = pd.DataFrame(st.session_state.log_report)
        st.dataframe(status_df, use_container_width=True)

    if not st.session_state.big_df.empty:
        df = st.session_state.big_df
        
        # 簡單過濾掉沒抓到的數據
        valid_df = df[df['學校'] != "未找到欄位"]
        
        if valid_df.empty:
            st.warning("⚠️ 雖然處理完成，但似乎沒有抓到任何有效數據。請檢查左側的欄位名稱是否與 Excel 完全一致。")
        else:
            st.subheader("📊 統計儀表板")
            
            # 指標
            m1, m2, m3 = st.columns(3)
            m1.metric("總參與人數", f"{len(valid_df)} 人")
            m2.metric("涵蓋縣市", f"{valid_df['縣市'].nunique()} 個")
            m3.metric("參與學校", f"{valid_df['學校'].nunique()} 所")
            
            # 圖表
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("#### 🌍 各縣市人數")
                st.bar_chart(valid_df['縣市'].value_counts())
            with c2:
                st.markdown("#### 🎓 職稱/身分比例")
                st.bar_chart(valid_df['職稱'].value_counts(), color="#ffaa00")

            # 下載按鈕
            st.subheader("📥 下載報告")
            if st.session_state.stats_excel:
                st.download_button(
                    label="📊 下載統計報表 (.xlsx)",
                    data=st.session_state.stats_excel,
                    file_name="自定義欄位_統計報表.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
