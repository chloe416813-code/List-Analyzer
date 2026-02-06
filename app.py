import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import datetime
import xlsxwriter
import openpyxl
from openpyxl.styles import PatternFill

# ================= 0. 系統設定 =================
st.set_page_config(page_title="科普列車統計系統 V7.0", page_icon="📊", layout="wide")

# 初始化 session state (這是解決下載按鈕消失的關鍵!)
if 'analysis_done' not in st.session_state:
    st.session_state.analysis_done = False
if 'result_zip' not in st.session_state:
    st.session_state.result_zip = None
if 'stats_excel' not in st.session_state:
    st.session_state.stats_excel = None
if 'big_df' not in st.session_state:
    st.session_state.big_df = pd.DataFrame()
if 'meta_report' not in st.session_state:
    st.session_state.meta_report = []

# ================= 1. 核心邏輯區 =================
REF_DATE = datetime(2025, 10, 20)

def parse_roc_birthday(roc_val):
    """ 解析民國年生日 """
    if pd.isna(roc_val): return None
    s = str(roc_val).strip().replace('\t', '').replace(' ', '')
    if s == '' or s.lower() == 'nan': return None
    s_clean = s.replace('年', '.').replace('月', '.').replace('日', '').replace('-', '.').replace('/', '.')
    
    parts = []
    if '.' in s_clean: parts = s_clean.split('.')
    elif s_clean.isdigit():
        if len(s_clean) == 6: parts = [s_clean[:2], s_clean[2:4], s_clean[4:]]
        elif len(s_clean) == 7: parts = [s_clean[:3], s_clean[3:5], s_clean[5:]]
    try:
        if len(parts) != 3: return None
        y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
        if not (1 <= m <= 12 and 1 <= d <= 31): return None
        return datetime(y + 1911, m, d)
    except:
        return None

def calculate_age(born):
    if born is None: return -1
    return REF_DATE.year - born.year - ((REF_DATE.month, REF_DATE.day) < (born.month, born.day))

def open_excel_safe(file_content, password):
    """ 安全開啟 Excel (支援讀取加密檔，但不輸出加密) """
    file_stream = io.BytesIO(file_content)
    try:
        return openpyxl.load_workbook(file_stream)
    except:
        file_stream.seek(0)
    
    # 如果直接開啟失敗，嘗試用密碼解鎖 (需安裝 msoffcrypto-tool)
    if password:
        try:
            import msoffcrypto
            decrypted = io.BytesIO()
            office_file = msoffcrypto.OfficeFile(file_stream)
            office_file.load_key(password=password)
            office_file.decrypt(decrypted)
            decrypted.seek(0)
            return openpyxl.load_workbook(decrypted)
        except:
            return None
    return None

def find_col_name(columns, keywords):
    """ 智慧搜尋欄位名稱 """
    for col in columns:
        col_str = str(col)
        if any(k in col_str for k in keywords):
            return col
    return None

def process_file_logic(filename, content, password):
    """ 核心處理：檢查黃底 + 萃取統計資料 """
    wb = open_excel_safe(content, password)
    if wb is None:
        return None, None, {"filename": filename, "status": "Fail", "msg": "無法開啟 (密碼錯誤或格式不支援)"}

    ws = wb.active
    
    # 1. 尋找表頭
    header_row_idx = 0
    # 掃描前 5 行找關鍵字
    for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True)):
        row_str = [str(c) if c else '' for c in row]
        if any('身分證' in s for s in row_str) or any('生日' in s for s in row_str):
            header_row_idx = r_idx
            break
    
    # 2. 準備統計用的 DataFrame
    data = list(ws.values)
    if not data: return None, None, {"filename": filename, "status": "Fail", "msg": "空檔案"}
    
    header = data[header_row_idx]
    rows = data[header_row_idx+1:]
    df = pd.DataFrame(rows, columns=header)
    
    # 3. 欄位對應
    cols = [str(c) for c in df.columns]
    
    # 定義關鍵字
    key_id = ['身分證', 'ID', '證號']
    key_birth = ['生日', '出生', 'Birth']
    key_city = ['縣市', '城市', 'City', '地區', '居住地']
    key_school = ['學校', '校名', 'School', '單位', '就讀學校']
    key_role = ['職稱', '身分', '身份', 'Role', '職務', '對象']

    col_id = find_col_name(cols, key_id)
    col_birth = find_col_name(cols, key_birth)
    col_city = find_col_name(cols, key_city)
    col_school = find_col_name(cols, key_school)
    col_role = find_col_name(cols, key_role)

    stats_meta = {"filename": filename, "under_15": 0, "adult": 0, "errors": 0, "status": "Success", "msg": "OK"}

    if not col_id or not col_birth:
        return None, None, {"filename": filename, "status": "Fail", "msg": "找不到關鍵欄位 (身分證/生日)"}

    # 4. 檢查邏輯 (標記黃底)
    # 為了寫入黃底，我們重新讀取 wb (因為上面只讀了值)
    wb_out = open_excel_safe(content, password)
    ws_out = wb_out.active
    YELLOW = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    
    # 找出 openpyxl 中的欄位 index
    op_header = list(ws_out.iter_rows(min_row=header_row_idx+1, max_row=header_row_idx+1, values_only=True))[0]
    op_header = [str(c) for c in op_header]
    
    try:
        idx_id = op_header.index(col_id)
        idx_birth = op_header.index(col_birth)
    except:
        return None, None, {"filename": filename, "status": "Fail", "msg": "欄位索引對應失敗"}

    for row in ws_out.iter_rows(min_row=header_row_idx+2):
        # 生日檢查
        if idx_birth < len(row):
            cell = row[idx_birth]
            dt = parse_roc_birthday(cell.value)
            if dt is None:
                cell.fill = YELLOW
                stats_meta["errors"] += 1
            else:
                age = calculate_age(dt)
                if 0 <= age < 15: stats_meta["under_15"] += 1
                elif age >= 15: stats_meta["adult"] += 1
        
        # 身分證檢查
        if idx_id < len(row):
            cell = row[idx_id]
            val = str(cell.value).strip() if cell.value else ""
            if not val or val == 'None' or len(val) != 10:
                cell.fill = YELLOW
                stats_meta["errors"] += 1

    output = io.BytesIO()
    wb_out.save(output)
    output.seek(0)
    
    # 5. 整理統計資料
    df_stat = df.copy()
    rename_map = {}
    if col_city: rename_map[col_city] = '縣市'
    else: df_stat['縣市'] = '未填縣市'
    
    if col_school: rename_map[col_school] = '學校'
    else: df_stat['學校'] = '未填學校'
    
    if col_role: rename_map[col_role] = '職稱'
    else: df_stat['職稱'] = '一般'
    
    df_stat.rename(columns=rename_map, inplace=True)
    
    # 確保欄位存在
    for c in ['縣市', '學校', '職稱']:
        if c not in df_stat.columns: df_stat[c] = 'Unknown'
            
    return output, df_stat, stats_meta

# ================= 2. 批量執行 =================

def run_analysis(files, pwd):
    processed_files = []
    report_list = []
    combined_df = pd.DataFrame() 
    
    progress_bar = st.progress(0)
    
    for i, f in enumerate(files):
        excel_data, df_stat, meta = process_file_logic(f.name, f.read(), pwd)
        report_list.append(meta)
        
        if excel_data:
            processed_files.append((f"已檢查_{f.name}", excel_data.getvalue()))
        
        if df_stat is not None:
            df_stat['來源檔案'] = f.name
            combined_df = pd.concat([combined_df, df_stat], ignore_index=True)
            
        progress_bar.progress((i + 1) / len(files))
        
    return processed_files, report_list, combined_df

# ================= 3. 主介面 =================

st.title("📊 科普列車 - 檢查與統計系統 V7.0")
st.markdown("---")

st.info("💡 說明：上傳 Excel 檔案後，系統會檢查格式（標記黃底）並自動產出 **Excel 統計報表**（各校師生人數）。")

col1, col2 = st.columns([1, 2])
with col1:
    pwd_input = st.text_input("檔案密碼 (若無則留空)", type="password", key="pwd_input")
with col2:
    files_input = st.file_uploader("請上傳 Excel (可多選合併計算)", type=['xlsx'], accept_multiple_files=True)

# 按下按鈕後，處理資料並存入 session_state
if st.button("🚀 開始分析 & 產生報表", type="primary"):
    if not files_input:
        st.warning("請先上傳檔案！")
    else:
        with st.spinner("正在分析數據中..."):
            res_files, meta_list, big_df = run_analysis(files_input, pwd_input)
            
            # 1. 處理檢查結果 ZIP
            if res_files:
                z = io.BytesIO()
                with zipfile.ZipFile(z, "w") as zf:
                    for n, d in res_files: zf.writestr(n, d)
                    txt = "\n".join([f"{r['filename']}: {r['msg']}" for r in meta_list])
                    zf.writestr("report.txt", txt)
                st.session_state.result_zip = z.getvalue()
            
            # 2. 處理統計報表 Excel
            if not big_df.empty:
                try:
                    stats_io = io.BytesIO()
                    # 製作樞紐分析表
                    pivot = big_df.pivot_table(index=['縣市', '學校'], columns='職稱', aggfunc='size', fill_value=0)
                    pivot['該校總計'] = pivot.sum(axis=1)
                    
                    with pd.ExcelWriter(stats_io, engine='xlsxwriter') as writer:
                        pivot.to_excel(writer, sheet_name='各校統計')
                        big_df['縣市'].value_counts().to_frame(name="人數").to_excel(writer, sheet_name='縣市統計')
                        big_df.to_excel(writer, sheet_name='總名單明細', index=False)
                        
                        # 美化欄寬
                        writer.sheets['各校統計'].set_column(0, 0, 15)
                        writer.sheets['各校統計'].set_column(1, 1, 30)
                        
                    st.session_state.stats_excel = stats_io.getvalue()
                    st.session_state.big_df = big_df
                except Exception as e:
                    st.error(f"統計報表產生失敗: {e}")
            
            st.session_state.meta_report = meta_list
            st.session_state.analysis_done = True

# ================= 4. 結果顯示區 =================
# 只有當分析完成後，這裡才會顯示。即使按了下載按鈕導致頁面重整，這裡的內容也會因為 session_state 而保留。

if st.session_state.analysis_done:
    st.divider()
    
    # --- 儀表板區 ---
    if not st.session_state.big_df.empty:
        df = st.session_state.big_df
        st.subheader("📈 數據儀表板")
        
        m1, m2, m3 = st.columns(3)
        m1.metric("總參與人數", f"{len(df)} 人")
        m2.metric("涵蓋縣市", f"{df['縣市'].nunique()} 個")
        m3.metric("參與學校", f"{df['學校'].nunique()} 所")
        
        c1, c2 = st.columns(2)
        with c1:
            st.caption("各縣市報名人數")
            st.bar_chart(df['縣市'].value_counts())
        with c2:
            st.caption("職稱比例")
            st.bar_chart(df['職稱'].value_counts(), color="#ffaa00")

    # --- 下載區 ---
    st.subheader("📥 下載報告")
    d1, d2 = st.columns(2)
    
    with d1:
