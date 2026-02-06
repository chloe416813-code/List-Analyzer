import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import datetime
import xlsxwriter
import openpyxl
from openpyxl.styles import PatternFill

# ================= 0. 系統設定 =================
st.set_page_config(page_title="科普列車統計系統 V8.1", page_icon="🚄", layout="wide")

# 初始化 Session State
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
    file_stream = io.BytesIO(file_content)
    try:
        return openpyxl.load_workbook(file_stream)
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
            return openpyxl.load_workbook(decrypted)
        except:
            return None
    return None

def get_single_col_name(columns, keywords):
    """ 嚴格篩選：只回傳【第一個】符合的欄位名稱 """
    for col in columns:
        col_str = str(col).strip()
        if col_str in keywords:
            return col
    for col in columns:
        col_str = str(col).strip()
        if any(k in col_str for k in keywords):
            return col
    return None

def process_file_logic(filename, content, password):
    wb = open_excel_safe(content, password)
    if wb is None:
        return None, None, {"filename": filename, "status": "Fail", "msg": "無法開啟 (密碼錯誤或格式不支援)"}

    ws = wb.active
    
    # --- A. 尋找表頭 ---
    header_row_idx = 0
    for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True)):
        row_str = [str(c) if c else '' for c in row]
        if any('身分證' in s for s in row_str) or any('生日' in s for s in row_str):
            header_row_idx = r_idx
            break
    
    # --- B. 讀取資料並【去重複】 ---
    data = list(ws.values)
    if not data: return None, None, {"filename": filename, "status": "Fail", "msg": "空檔案"}
    
    raw_header = data[header_row_idx]
    rows = data[header_row_idx+1:]
    
    # 建立 DataFrame
    df = pd.DataFrame(rows, columns=raw_header)
    
    # === 關鍵修正：處理重複欄位名稱 ===
    # 如果 Excel 有兩個「職稱」欄位，pandas 會造成 Grouper error
    # 我們這裡強制重新命名重複的欄位
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated()] # 只保留第一個出現的欄位
    
    cols = list(df.columns)
    
    # --- C. 關鍵字定義 ---
    key_id = ['身分證', 'ID', '證號', '身分證字號']
    key_birth = ['生日', '出生', 'Birth', '出生年月日']
    key_city = ['縣市', '城市', 'City', '地區', '居住地', '縣市別']
    key_school = ['學校', '校名', 'School', '單位', '就讀學校', '學校名稱']
    key_role = ['職稱', '身分', '身份', 'Role', '職務', '對象', '類別', '師生']

    col_id = get_single_col_name(cols, key_id)
    col_birth = get_single_col_name(cols, key_birth)
    col_city = get_single_col_name(cols, key_city)
    col_school = get_single_col_name(cols, key_school)
    col_role = get_single_col_name(cols, key_role)

    stats_meta = {"filename": filename, "under_15": 0, "adult": 0, "errors": 0, "status": "Success", "msg": "OK"}

    if not col_id or not col_birth:
        return None, None, {"filename": filename, "status": "Fail", "msg": "找不到關鍵欄位 (身分證/生日)"}

    # --- D. 黃底檢查 (僅標記) ---
    wb_out = open_excel_safe(content, password)
    ws_out = wb_out.active
    YELLOW = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    
    # 重新對應 index (針對 openpyxl)
    op_header_row = list(ws_out.iter_rows(min_row=header_row_idx+1, max_row=header_row_idx+1, values_only=True))[0]
    op_header_str = [str(c).strip() for c in op_header_row]
    
    try:
        # 找出 openpyxl 對應的 index (只找第一個匹配的)
        idx_id = op_header_str.index(col_id)
        idx_birth = op_header_str.index(col_birth)
        
        for row in ws_out.iter_rows(min_row=header_row_idx+2):
            # 生日
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
            # 身分證
            if idx_id < len(row):
                cell = row[idx_id]
                val = str(cell.value).strip() if cell.value else ""
                if not val or val == 'None' or len(val) != 10:
                    cell.fill = YELLOW
                    stats_meta["errors"] += 1
    except:
        stats_meta["msg"] = "檢查過程發生索引警告 (不影響統計)"

    output = io.BytesIO()
    wb_out.save(output)
    output.seek(0)
    
    # --- E. 萃取純淨統計資料 ---
    clean_data = {}
    
    clean_data['縣市'] = df[col_city] if col_city else '未填縣市'
    clean_data['學校'] = df[col_school] if col_school else '未填學校'
    
    if col_role:
        clean_data['職稱'] = df[col_role]
    else:
        # 若無職稱欄位，用生日推算
        def guess_role(row):
            dt = parse_roc_birthday(row.get(col_birth))
            if dt:
                return '學生' if calculate_age(dt) < 15 else '師長/成人'
            return '一般'
        clean_data['職稱'] = df.apply(guess_role, axis=1)

    df_stat = pd.DataFrame(clean_data)
    df_stat.fillna("未知", inplace=True)
    
    return output, df_stat, stats_meta

# ================= 2. 批量執行 =================

def run_analysis(files, pwd):
    processed_files = []
    report_list = []
    combined_df = pd.DataFrame() 
    
    progress_bar = st.progress(0)
    
    for i, f in enumerate(files):
        try:
            excel_data, df_stat, meta = process_file_logic(f.name, f.read(), pwd)
            report_list.append(meta)
            
            if excel_data:
                processed_files.append((f"已檢查_{f.name}", excel_data.getvalue()))
            
            if df_stat is not None:
                df_stat['來源檔案'] = f.name
                combined_df = pd.concat([combined_df, df_stat], ignore_index=True)
        except Exception as e:
            st.error(f"檔案 {f.name} 處理失敗: {e}")
            
        progress_bar.progress((i + 1) / len(files))
        
    return processed_files, report_list, combined_df

# ================= 3. 主介面 =================

st.title("🚄 科普列車 - 統計戰情室 V8.1")
st.markdown("### 專注於：縣市、學校、師生人數統計")
st.info("已修正 Grouper error，系統將自動過濾重複欄位。")

col1, col2 = st.columns([1, 2])
with col1:
    pwd_input = st.text_input("檔案密碼 (若無則留空)", type="password", key="pwd_input")
with col2:
    files_input = st.file_uploader("請上傳 Excel (支援多檔合併)", type=['xlsx'], accept_multiple_files=True)

if st.button("🚀 開始分析 & 產生報表", type="primary"):
    if not files_input:
        st.warning("請先上傳檔案！")
    else:
        # 清除舊狀態
        st.session_state.big_df = pd.DataFrame()
        
        with st.spinner("正在進行智慧欄位辨識與統計..."):
            res_files, meta_list, big_df = run_analysis(files_input, pwd_input)
            
            # 存 Session State
            if res_files:
                z = io.BytesIO()
                with zipfile.ZipFile(z, "w") as zf:
                    for n, d in res_files: zf.writestr(n, d)
                    txt = "\n".join([f"{r['filename']}: {r['msg']}" for r in meta_list])
                    zf.writestr("report.txt", txt)
                st.session_state.result_zip = z.getvalue()
            
            if not big_df.empty:
                try:
                    stats_io = io.BytesIO()
                    # 樞紐分析
                    pivot = big_df.pivot_table(index=['縣市', '學校'], columns='職稱', aggfunc='size', fill_value=0)
                    pivot['該校總計'] = pivot.sum(axis=1)
                    
                    with pd.ExcelWriter(stats_io, engine='xlsxwriter') as writer:
                        pivot.to_excel(writer, sheet_name='各校統計')
                        big_df['縣市'].value_counts().to_frame(name="人數").to_excel(writer, sheet_name='縣市統計')
                        big_df.to_excel(writer, sheet_name='總名單明細', index=False)
                        writer.sheets['各校統計'].set_column(0, 1, 20)
                        
                    st.session_state.stats_excel = stats_io.getvalue()
                    st.session_state.big_df = big_df
                    
                except Exception as e:
                    st.error(f"統計報表產生失敗: {e}")
            
            st.session_state.meta_report = meta_list
            st.session_state.analysis_done = True

# ================= 4. 結果顯示區 =================

if st.session_state.analysis_done:
    st.divider()
    
    if not st.session_state.big_df.empty:
        df = st.session_state.big_df
        st.subheader("📊 統計儀表板")
        
        m1, m2, m3 = st.columns(3)
        m1.metric("總參與人數", f"{len(df)} 人")
        m2.metric("涵蓋縣市", f"{df['縣市'].nunique()} 個")
        m3.metric("參與學校", f"{df['學校'].nunique()} 所")
        
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### 🌍 各縣市人數")
            st.bar_chart(df['縣市'].value_counts())
        with c2:
            st.markdown("#### 🎓 師生職稱比例")
            st.bar_chart(df['職稱'].value_counts(), color="#ffaa00")

    st.subheader("📥 下載報告")
    d1, d2 = st.columns(2)
    
    with d1:
        if st.session_state.stats_excel:
            st.download_button(
                label="📊 下載統計報表 (.xlsx)",
                data=st.session_state.stats_excel,
                file_name="科普列車_統計報表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
    with d2:
        if st.session_state.result_zip:
            st.download_button(
                label="📦 下載已檢查原始檔 (ZIP)",
                data=st.session_state.result_zip,
                file_name="檢查結果_黃底標記.zip",
                mime="application/zip"
            )
            
    with st.expander("查看檢查詳細日誌"):
        st.dataframe(pd.DataFrame(st.session_state.meta_report))
