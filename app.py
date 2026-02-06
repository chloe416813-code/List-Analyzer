import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import datetime
import xlsxwriter
import openpyxl
from openpyxl.styles import PatternFill

# ================= 0. 系統設定 =================
st.set_page_config(page_title="科普列車統計系統 V9.0", page_icon="🚄", layout="wide")

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

def find_best_column(columns, keywords):
    """
    V9.0 核心演算法：模糊搜尋欄位
    只要欄位名稱【包含】關鍵字，就抓取。
    """
    # 1. 先找完全一樣的 (優先權最高)
    for col in columns:
        col_str = str(col).strip()
        if col_str in keywords:
            return col
            
    # 2. 再找包含關鍵字的 (例如 keywords=['學校']，可以抓到 '就讀學校')
    for col in columns:
        col_str = str(col).strip()
        for k in keywords:
            if k in col_str:
                return col
    return None

def process_file_logic(filename, content, password):
    """ 核心處理 """
    wb = open_excel_safe(content, password)
    if wb is None:
        return None, None, {"filename": filename, "status": "Fail", "msg": "無法開啟(密碼錯誤)"}

    ws = wb.active
    
    # A. 找表頭
    header_row_idx = 0
    # 掃描前 5 行，只要有任何一格有值，我們就猜測這行可能是表頭的開始
    # 這裡放寬標準，因為使用者的格式可能很多元
    for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True)):
        row_str = [str(c) for c in row if c is not None]
        # 如果這一行包含我們關注的關鍵字，那它肯定是表頭
        full_str = "".join(row_str)
        if "學校" in full_str or "縣市" in full_str or "身分" in full_str or "生日" in full_str:
            header_row_idx = r_idx
            break
    
    # B. 讀取資料
    data = list(ws.values)
    if not data: return None, None, {"filename": filename, "status": "Fail", "msg": "空檔案"}
    
    raw_header = data[header_row_idx]
    rows = data[header_row_idx+1:]
    
    df = pd.DataFrame(rows, columns=raw_header)
    
    # === 去除重複欄位，避免報錯 ===
    df.columns = [str(c).strip() if c is not None else f"Unnamed_{i}" for i, c in enumerate(df.columns)]
    df = df.loc[:, ~df.columns.duplicated()]
    
    cols = list(df.columns)
    
    # C. 定義關鍵字 (越前面越優先)
    # 我們把所有可能的叫法都寫進去
    key_city = ['縣市', '城市', 'City', '地區', '居住地', '縣市別', '地址', '住址']
    key_school = ['學校', '校名', 'School', '單位', '就讀學校', '學校名稱', '服務單位']
    # 這裡的 "職稱" 我們用來區分 老師/學生
    key_role = ['職稱', '身分', '身份', 'Role', '職務', '對象', '類別', '師生']
    
    # 檢查用 (若使用者還是希望有檢查功能)
    key_birth = ['生日', '出生', 'Birth']

    # 自動抓取對應欄位
    col_city = find_best_column(cols, key_city)
    col_school = find_best_column(cols, key_school)
    col_role = find_best_column(cols, key_role)
    col_birth = find_best_column(cols, key_birth)

    stats_meta = {"filename": filename, "msg": "OK"}

    # D. 萃取純淨統計資料
    clean_data = {}
    
    # 1. 抓縣市
    if col_city:
        clean_data['縣市'] = df[col_city]
    else:
        clean_data['縣市'] = '未偵測到縣市'

    # 2. 抓學校
    if col_school:
        clean_data['學校'] = df[col_school]
    else:
        clean_data['學校'] = '未偵測到學校'
        
    # 3. 抓身分/職稱 (判斷人數類別)
    if col_role:
        clean_data['職稱'] = df[col_role]
    else:
        # 如果沒職稱欄位，嘗試用生日推算
        if col_birth:
            def guess_role(row):
                dt = parse_roc_birthday(row.get(col_birth))
                if dt: return '學生' if calculate_age(dt) < 15 else '師長/成人'
                return '一般'
            clean_data['職稱'] = df.apply(guess_role, axis=1)
        else:
            clean_data['職稱'] = '一般人員' # 真的都沒有，就統稱一般人員

    # 建立統計用 DataFrame
    df_stat = pd.DataFrame(clean_data)
    df_stat.fillna("未知", inplace=True)
    
    # 回傳：原始Excel(這邊不回傳檢查版以節省資源), 統計表, 訊息
    return None, df_stat, stats_meta

# ================= 2. 批量執行 =================

def run_analysis(files, pwd):
    report_list = []
    combined_df = pd.DataFrame() 
    
    progress_bar = st.progress(0)
    
    for i, f in enumerate(files):
        try:
            _, df_stat, meta = process_file_logic(f.name, f.read(), pwd)
            report_list.append(meta)
            
            if df_stat is not None:
                df_stat['來源檔案'] = f.name
                combined_df = pd.concat([combined_df, df_stat], ignore_index=True)
        except Exception as e:
            st.error(f"檔案 {f.name} 處理失敗: {e}")
            
        progress_bar.progress((i + 1) / len(files))
        
    return report_list, combined_df

# ================= 3. 主介面 =================

st.title("🚄 科普列車 - 智慧統計看板 V9.0")
st.markdown("### 自動抓取：[縣市]、[學校]、[人數] (依列計算)")
st.info("💡 只要 Excel 欄位名稱包含「縣市」、「學校」等關鍵字，系統就會自動吸附並進行統計。")

col1, col2 = st.columns([1, 2])
with col1:
    pwd_input = st.text_input("檔案密碼 (若無則留空)", type="password", key="pwd_input")
with col2:
    files_input = st.file_uploader("請上傳 Excel (支援多檔合併)", type=['xlsx'], accept_multiple_files=True)

if st.button("🚀 開始分析 & 產生報表", type="primary"):
    if not files_input:
        st.warning("請先上傳檔案！")
    else:
        st.session_state.big_df = pd.DataFrame()
        
        with st.spinner("正在進行智慧關鍵字分析..."):
            meta_list, big_df = run_analysis(files_input, pwd_input)
            
            if not big_df.empty:
                try:
                    stats_io = io.BytesIO()
                    
                    # === 產生統計報表 ===
                    # 1. 各校統計 (樞紐分析)
                    # 邏輯：列出 [縣市][學校]，算出 [職稱] 的人數
                    pivot = big_df.pivot_table(index=['縣市', '學校'], columns='職稱', aggfunc='size', fill_value=0)
                    pivot['該校總計'] = pivot.sum(axis=1)
                    
                    # 2. 縣市統計
                    city_counts = big_df['縣市'].value_counts().to_frame(name="人數")
                    
                    # 寫入 Excel
                    with pd.ExcelWriter(stats_io, engine='xlsxwriter') as writer:
                        pivot.to_excel(writer, sheet_name='各校師生統計')
                        city_counts.to_excel(writer, sheet_name='縣市統計')
                        big_df.to_excel(writer, sheet_name='總名單明細', index=False)
                        
                        # 美化
                        writer.sheets['各校師生統計'].set_column(0, 1, 20)
                        
                    st.session_state.stats_excel = stats_io.getvalue()
                    st.session_state.big_df = big_df
                    
                except Exception as e:
                    st.error(f"報表產生失敗: {e}")
            
            st.session_state.meta_report = meta_list
            st.session_state.analysis_done = True

# ================= 4. 結果顯示區 =================

if st.session_state.analysis_done:
    st.divider()
    
    if not st.session_state.big_df.empty:
        df = st.session_state.big_df
        st.subheader("📊 統計儀表板")
        
        # 指標
        m1, m2, m3 = st.columns(3)
        m1.metric("總參與人數", f"{len(df)} 人")
        m2.metric("涵蓋縣市", f"{df['縣市'].nunique()} 個")
        m3.metric("參與學校", f"{df['學校'].nunique()} 所")
        
        # 圖表
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### 🌍 各縣市人數分佈")
            st.bar_chart(df['縣市'].value_counts())
        with c2:
            st.markdown("#### 🎓 師生/職稱 比例")
            st.bar_chart(df['職稱'].value_counts(), color="#ffaa00")

    # 下載區
    st.subheader("📥 下載報告")
    
    if st.session_state.stats_excel:
        st.download_button(
            label="📊 下載統計報表 (.xlsx)",
            data=st.session_state.stats_excel,
            file_name="科普列車_統計報表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        st.caption("檔案內容包含：各校師生人數詳表、縣市人數統計、合併後總名單。")
            
    with st.expander("查看處理狀態"):
        st.dataframe(pd.DataFrame(st.session_state.meta_report))
