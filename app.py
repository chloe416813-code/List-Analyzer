import streamlit as st
import pandas as pd
import io
import xlsxwriter
import openpyxl

# ================= 0. 系統設定 =================
st.set_page_config(page_title="科普列車統計系統 V13.0", page_icon="🚄", layout="wide")

# 初始化 Session State
if 'merged_df' not in st.session_state:
    st.session_state.merged_df = pd.DataFrame()
if 'step' not in st.session_state:
    st.session_state.step = 1  # 1:上傳, 2:確認身分, 3:結果

# ================= 1. 核心邏輯：自動抓取 =================

def open_excel_safe(file_content, password):
    """ 安全開啟 Excel """
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

def find_best_column(columns, keywords):
    """ 智慧模糊搜尋：只要欄位名稱包含關鍵字就回傳 """
    for col in columns:
        col_str = str(col).strip()
        for k in keywords:
            if k in col_str:
                return col
    return None

def extract_essential_data(filename, content, password):
    """ 
    核心功能：只抓 [縣市, 學校, 身分]，其他通通不要 
    """
    wb = open_excel_safe(content, password)
    if wb is None: return None, "無法讀取(密碼錯誤?)"
    ws = wb.active
    
    # 1. 自動找表頭 (掃描前10行)
    header_idx = 0
    found = False
    for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True)):
        row_str = "".join([str(c) for c in row if c is not None])
        # 只要這行同時出現 "學校" 或 "縣市"，就認定是表頭
        if "學校" in row_str or "縣市" in row_str or "地區" in row_str:
            header_idx = r_idx
            found = True
            break
    
    if not found: return None, "找不到關鍵表頭(需包含'學校'或'縣市')"

    # 2. 讀取資料
    data = list(ws.values)
    raw_header = data[header_idx]
    rows = data[header_idx+1:]
    
    df = pd.DataFrame(rows, columns=raw_header)
    # 清洗欄位名
    df.columns = [str(c).strip() for c in df.columns]
    cols = list(df.columns)

    # 3. 定義關鍵字庫 (自動吸附用)
    keys_city = ['縣市', '城市', '地區', '居住', 'City', '地點']
    keys_school = ['學校', '校名', '單位', 'School', '就讀']
    keys_role = ['職稱', '身分', '身份', 'Role', '職務', '對象', '師生']

    # 4. 抓取對應欄位
    col_city = find_best_column(cols, keys_city)
    col_school = find_best_column(cols, keys_school)
    col_role = find_best_column(cols, keys_role)

    # 5. 建立純淨資料表 (只留這三欄)
    clean_data = {}
    
    if col_city: clean_data['縣市'] = df[col_city]
    else: clean_data['縣市'] = '未偵測到'
        
    if col_school: clean_data['學校'] = df[col_school]
    else: clean_data['學校'] = '未偵測到'
        
    if col_role: clean_data['職稱'] = df[col_role]
    else: clean_data['職稱'] = '未偵測到' # 如果沒身分欄，後續就無法分師生，但能算總數

    new_df = pd.DataFrame(clean_data)
    new_df.fillna("未知", inplace=True)
    
    return new_df, "OK"

# ================= 2. 主介面 =================

st.title("🚄 科普列車 - 自動統計戰情室 V13.0")
st.markdown("### 全自動模式：只抓取【縣市、學校、職稱】，忽略其他雜訊。")

# --- 側邊欄 ---
with st.sidebar:
    st.header("設定")
    pwd_input = st.text_input("檔案密碼 (若無則留空)", type="password")
    if st.button("🔄 重置所有分析"):
        st.session_state.merged_df = pd.DataFrame()
        st.session_state.step = 1
        st.rerun()

# --- 步驟 1: 上傳與自動抓取 ---
if st.session_state.step == 1:
    st.info("請上傳 Excel 檔案，系統會自動尋找關鍵欄位進行合併。")
    files = st.file_uploader("上傳 Excel (支援多檔)", type=['xlsx'], accept_multiple_files=True)
    
    if st.button("🚀 開始自動抓取", type="primary"):
        if files:
            all_dfs = []
            bar = st.progress(0)
            log = []
            
            for i, f in enumerate(files):
                df, msg = extract_essential_data(f.name, f.read(), pwd_input)
                if df is not None:
                    df['來源檔案'] = f.name
                    all_dfs.append(df)
                    log.append(f"✅ {f.name}: 成功")
                else:
                    log.append(f"❌ {f.name}: {msg}")
                bar.progress((i+1)/len(files))
            
            if all_dfs:
                st.session_state.merged_df = pd.concat(all_dfs, ignore_index=True)
                # 顯示處理結果
                with st.expander("查看讀取狀態"):
                    st.write(log)
                st.session_state.step = 2
                st.rerun()
            else:
                st.error("沒有成功讀取任何檔案，請確認 Excel 包含「縣市」或「學校」相關欄位。")
        else:
            st.warning("請先上傳檔案")

# --- 步驟 2: 智慧分類 (師生辨識) ---
if st.session_state.step == 2:
    st.header("步驟 2: 確認身分分類")
    
    df = st.session_state.merged_df
    # 取得所有職稱
    unique_roles = [str(r) for r in df['職稱'].unique() if r != '未知']
    
    # 自動猜測 (Auto-Guess Logic)
    # 只要字串裡有 "師" 或 "長"，就預設勾選為老師
    default_teachers = [r for r in unique_roles if any(x in r for x in ['師', '長', '教', '授'])]
    # 只要字串裡有 "生"，就預設勾選為學生
    default_students = [r for r in unique_roles if '生' in r]

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("🧑‍🏫 老師/成人")
        teachers = st.multiselect("請確認或修改", unique_roles, default=default_teachers, key="sel_T")
    with c2:
        st.subheader("👶 學生")
        students = st.multiselect("請確認或修改", unique_roles, default=default_students, key="sel_S")
        
    st.caption("未被勾選的職稱將歸類為「其他」，但仍會計入總人數。")

    if st.button("📊 生成儀表板與報表", type="primary"):
        # 標記身分
        def tag_role(r):
            r = str(r)
            if r in teachers: return '老師'
            if r in students: return '學生'
            return '其他'
        
        df['類別'] = df['職稱'].apply(tag_role)
        st.session_state.merged_df = df # 更新 DataFrame
        st.session_state.step = 3
        st.rerun()

# --- 步驟 3: 儀表板與下載 ---
if st.session_state.step == 3:
    df = st.session_state.merged_df
    
    st.divider()
    
    # 1. 關鍵指標
    st.subheader("📈 統計儀表板")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("總參與人數", f"{len(df)} 人")
    m2.metric("涵蓋縣市", f"{df['縣市'].nunique()} 個")
    m3.metric("參與學校", f"{df['學校'].nunique()} 所")
    
    # 計算師生比
    t_count = len(df[df['類別']=='老師'])
    s_count = len(df[df['類別']=='學生'])
    m4.metric("師生結構", f"師{t_count} : 生{s_count}")

    # 2. 圖表
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**各縣市參與人數**")
        st.bar_chart(df['縣市'].value_counts())
    with c2:
        st.markdown("**身分比例**")
        st.bar_chart(df['類別'].value_counts(), color="#ffaa00")

    # 3. 產生統計表 (Pivot Table)
    # 我們要做一張表：[縣市] [學校] [老師] [學生] [其他] [總計]
    pivot = df.pivot_table(index=['縣市', '學校'], columns='類別', aggfunc='size', fill_value=0)
    
    # 確保欄位存在 (防止某個類別完全沒人)
    for col in ['老師', '學生', '其他']:
        if col not in pivot.columns: pivot[col] = 0
            
    # 計算總計
    pivot['該校總計'] = pivot['老師'] + pivot['學生'] + pivot['其他']
    
    # 調整欄位順序
    pivot = pivot[['老師', '學生', '其他', '該校總計']]

    st.subheader("🏫 各校統計明細")
    st.dataframe(pivot, use_container_width=True)

    # 4. 下載
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        pivot.to_excel(writer, sheet_name='各校統計表')
        
        # 縣市總表
        city_stats = df.pivot_table(index='縣市', columns='類別', aggfunc='size', fill_value=0)
        city_stats['總計'] = city_stats.sum(axis=1)
        city_stats.to_excel(writer, sheet_name='縣市統計表')
        
        # 原始清洗後名單
        df.to_excel(writer, sheet_name='總名單明細', index=False)
        
        # 美化
        writer.sheets['各校統計表'].set_column(0, 1, 20)
        
    st.download_button(
        label="📥 下載完整統計報表 (Excel)",
        data=output.getvalue(),
        file_name="科普列車_統計分析.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
