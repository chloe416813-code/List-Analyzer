import streamlit as st
import pandas as pd
import io
import xlsxwriter
import openpyxl

# ================= 0. 系統設定 =================
st.set_page_config(page_title="科普列車統計系統 V12.0", page_icon="📊", layout="wide")

# 初始化 Session State
if 'merged_df' not in st.session_state:
    st.session_state.merged_df = pd.DataFrame()
if 'step' not in st.session_state:
    st.session_state.step = 1 # 1:上傳, 2:設定職稱, 3:看結果

# ================= 1. 核心邏輯區 =================

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

def extract_data(filename, content, password, user_cols_map):
    """
    抓取資料並標準化欄位名稱
    user_cols_map = {'city': '居住縣市', 'school': '就讀學校', 'role': '身分'}
    """
    wb = open_excel_safe(content, password)
    if wb is None: return None
    ws = wb.active
    
    # 找表頭
    header_idx = 0
    found = False
    # 搜尋前 10 行，看哪一行包含了使用者設定的「學校」欄位名稱
    target_key = user_cols_map['school'].strip()
    
    for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True)):
        row_str = [str(c).strip() for c in row if c is not None]
        if target_key in row_str:
            header_idx = r_idx
            found = True
            break
            
    if not found: return None # 找不到表頭

    # 讀取資料
    data = list(ws.values)
    raw_header = data[header_idx]
    rows = data[header_idx+1:]
    
    df = pd.DataFrame(rows, columns=raw_header)
    df.columns = [str(c).strip() for c in df.columns] # 清洗欄位名
    
    # 萃取並重新命名欄位 (標準化)
    # 我們將使用者 Excel 裡千奇百怪的欄位名，統一改成 'City', 'School', 'Role'
    clean_data = {}
    
    # 對應 City
    col_city = user_cols_map['city'].strip()
    if col_city in df.columns: clean_data['City'] = df[col_city]
    else: clean_data['City'] = '未知縣市'
        
    # 對應 School
    col_school = user_cols_map['school'].strip()
    if col_school in df.columns: clean_data['School'] = df[col_school]
    else: clean_data['School'] = '未知學校'
        
    # 對應 Role
    col_role = user_cols_map['role'].strip()
    if col_role in df.columns: clean_data['Role'] = df[col_role]
    else: clean_data['Role'] = '未知'

    return pd.DataFrame(clean_data)

# ================= 2. 主介面 =================

st.title("📊 科普列車 - 縣市學校統計戰情室 V12.0")
st.markdown("### 專注目標：算出各縣市有【幾間學校】、【幾位老師】、【幾位學生】")

# --- 左側設定區 ---
with st.sidebar:
    st.header("1. 欄位對應設定")
    st.info("請告訴系統，您的 Excel 裡這些欄位叫什麼？")
    
    input_city = st.text_input("縣市欄位名稱", value="居住縣市")
    input_school = st.text_input("學校欄位名稱", value="就讀學校")
    input_role = st.text_input("身分/職稱欄位名稱", value="職稱")
    
    st.divider()
    pwd_input = st.text_input("檔案密碼 (若無則留空)", type="password")
    
    # 打包設定
    cols_map = {'city': input_city, 'school': input_school, 'role': input_role}

# --- 步驟 1: 上傳與讀取 ---
if st.session_state.step == 1:
    st.header("步驟 1: 上傳資料")
    files = st.file_uploader("上傳 Excel (支援多檔)", type=['xlsx'], accept_multiple_files=True)
    
    if st.button("讀取資料 & 下一步", type="primary"):
        if files:
            all_dfs = []
            bar = st.progress(0)
            for i, f in enumerate(files):
                df = extract_data(f.name, f.read(), pwd_input, cols_map)
                if df is not None:
                    df['Source'] = f.name
                    all_dfs.append(df)
                bar.progress((i+1)/len(files))
            
            if all_dfs:
                st.session_state.merged_df = pd.concat(all_dfs, ignore_index=True)
                # 填充空值
                st.session_state.merged_df.fillna("未知", inplace=True)
                st.session_state.step = 2
                st.rerun() # 重新整理進入下一步
            else:
                st.error("讀取失敗，請檢查欄位名稱設定是否正確。")
        else:
            st.warning("請先上傳檔案")

# --- 步驟 2: 設定統計邏輯 (這是解決您問題的關鍵) ---
if st.session_state.step == 2:
    st.header("步驟 2: 定義身分 (誰是老師？誰是學生？)")
    
    df = st.session_state.merged_df
    
    # 抓出 Excel 裡所有出現過的職稱
    unique_roles = df['Role'].unique().tolist()
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("🧑‍🏫 哪些職稱算「老師/成人」？")
        # 讓使用者多選
        teacher_list = st.multiselect("請勾選 (可多選)", unique_roles, default=[r for r in unique_roles if '師' in str(r) or '長' in str(r)])
        
    with col2:
        st.subheader("👶 哪些職稱算「學生」？")
        student_list = st.multiselect("請勾選 (可多選)", unique_roles, default=[r for r in unique_roles if '生' in str(r)])

    st.info(f"目前資料共 {len(df)} 筆。未被勾選的職稱將不計入師生統計，但會計入總人數。")

    if st.button("開始計算統計報表 🚀", type="primary"):
        # 開始進行聚合分析 (Aggregation)
        
        # 1. 先把每一列標記是老師還是學生
        def tag_role(role_val):
            if role_val in teacher_list: return 'Teacher'
            if role_val in student_list: return 'Student'
            return 'Other'
            
        df['Tag'] = df['Role'].apply(tag_role)
        
        # 2. 進行 GroupBy 統計
        # 我們要對 [City] 進行分組
        # 計算 School 的 nunique (不重複數量)
        # 計算 Tag 的 count
        
        summary_list = []
        
        # 針對每個縣市進行迴圈統計
        for city, group in df.groupby('City'):
            # 該縣市有多少間不重複的學校
            school_count = group['School'].nunique()
            
            # 該縣市有多少老師
            teacher_count = len(group[group['Tag'] == 'Teacher'])
            
            # 該縣市有多少學生
            student_count = len(group[group['Tag'] == 'Student'])
            
            # 該縣市總人數 (包含未分類的)
            total_count = len(group)
            
            summary_list.append({
                '縣市': city,
                '學校數': school_count,
                '老師人數': teacher_count,
                '學生人數': student_count,
                '總參與人數': total_count
            })
            
        # 轉成 DataFrame
        summary_df = pd.DataFrame(summary_list)
        
        # 排序 (依總人數多到少)
        summary_df = summary_df.sort_values(by='總參與人數', ascending=False).reset_index(drop=True)
        
        # 存入 Session 供下載
        st.session_state.summary_df = summary_df
        
        # 另外做一張「各校明細表」
        # Group by City AND School
        school_list = []
        for (city, school), group in df.groupby(['City', 'School']):
            t_c = len(group[group['Tag'] == 'Teacher'])
            s_c = len(group[group['Tag'] == 'Student'])
            school_list.append({
                '縣市': city,
                '學校': school,
                '老師人數': t_c,
                '學生人數': s_c,
                '該校總計': len(group)
            })
        st.session_state.school_df = pd.DataFrame(school_list)
        
        st.session_state.step = 3
        st.rerun()

# --- 步驟 3: 顯示結果與下載 ---
if st.session_state.step == 3:
    st.header("步驟 3: 分析結果")
    
    summary_df = st.session_state.summary_df
    school_df = st.session_state.school_df
    
    # 1. 顯示總表 (這就是您要的！)
    st.subheader("🏆 縣市統計總表")
    st.dataframe(summary_df, use_container_width=True)
    
    # 2. 顯示圖表
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("##### 各縣市學校數")
        st.bar_chart(summary_df.set_index('縣市')['學校數'])
    with c2:
        st.markdown("##### 各縣市參與人數")
        st.bar_chart(summary_df.set_index('縣市')['總參與人數'])

    # 3. 下載
    st.divider()
    
    # 製作 Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        summary_df.to_excel(writer, sheet_name='縣市統計總表', index=False)
        school_df.to_excel(writer, sheet_name='各校詳細統計', index=False)
        st.session_state.merged_df.to_excel(writer, sheet_name='原始合併名單', index=False)
        
        # 美化
        writer.sheets['縣市統計總表'].set_column(0, 4, 15)
        writer.sheets['各校詳細統計'].set_column(0, 1, 20)
        
    st.download_button(
        label="📥 下載完整統計報表 (Excel)",
        data=output.getvalue(),
        file_name="科普列車_統計分析.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
    
    if st.button("🔄 重新分析其他檔案"):
        st.session_state.step = 1
        st.session_state.merged_df = pd.DataFrame()
        st.rerun()
