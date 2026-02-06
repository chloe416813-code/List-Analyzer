import streamlit as st
import pandas as pd
import io
import xlsxwriter
import openpyxl

# ================= 0. 系統設定 =================
st.set_page_config(page_title="科普列車統計系統 V12.1", page_icon="📊", layout="wide")

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
    """
    wb = open_excel_safe(content, password)
    if wb is None: return None
    ws = wb.active
    
    # 找表頭
    header_idx = 0
    found = False
    target_key = str(user_cols_map['school']).strip()
    
    # 搜尋前 20 行 (放寬搜尋範圍)
    for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=20, values_only=True)):
        row_str = [str(c).strip() for c in row if c is not None]
        if target_key in row_str:
            header_idx = r_idx
            found = True
            break
            
    if not found: return None # 找不到表頭

    # 讀取資料
    data = list(ws.values)
    if not data: return None

    raw_header = data[header_idx]
    rows = data[header_idx+1:]
    
    df = pd.DataFrame(rows, columns=raw_header)
    # 清洗欄位名 (轉字串、去空白)
    df.columns = [str(c).strip() for c in df.columns]
    
    # 萃取並重新命名欄位 (標準化)
    clean_data = {}
    
    # 對應 City
    col_city = str(user_cols_map['city']).strip()
    if col_city in df.columns: 
        clean_data['City'] = df[col_city]
    else: 
        clean_data['City'] = '未知縣市'
        
    # 對應 School
    col_school = str(user_cols_map['school']).strip()
    if col_school in df.columns: 
        clean_data['School'] = df[col_school]
    else: 
        clean_data['School'] = '未知學校'
        
    # 對應 Role
    col_role = str(user_cols_map['role']).strip()
    if col_role in df.columns: 
        clean_data['Role'] = df[col_role]
    else: 
        clean_data['Role'] = '未知'

    return pd.DataFrame(clean_data)

# ================= 2. 主介面 =================

st.title("📊 科普列車 - 縣市學校統計戰情室 V12.1")
st.markdown("### 專注目標：算出各縣市有【幾間學校】、【幾位老師】、【幾位學生】")

# --- 左側設定區 ---
with st.sidebar:
    st.header("1. 欄位對應設定")
    st.info("請輸入 Excel 裡對應的表頭名稱 (完全一致)")
    
    input_city = st.text_input("縣市欄位名稱", value="請選擇學校所屬縣市")
    input_school = st.text_input("學校欄位名稱", value="參加學校")
    input_role = st.text_input("身分/職稱欄位名稱", value="參與師生總人數")
    
    st.divider()
    pwd_input = st.text_input("檔案密碼 (若無則留空)", type="password")
    
    # 打包設定
    cols_map = {'city': input_city, 'school': input_school, 'role': input_role}
    
    if st.button("🔄 重置所有步驟", type="secondary"):
        st.session_state.merged_df = pd.DataFrame()
        st.session_state.step = 1
        st.rerun()

# --- 步驟 1: 上傳與讀取 ---
if st.session_state.step == 1:
    st.header("步驟 1: 上傳資料")
    files = st.file_uploader("上傳 Excel (支援多檔)", type=['xlsx'], accept_multiple_files=True)
    
    if st.button("讀取資料 & 下一步", type="primary"):
        if files:
            all_dfs = []
            bar = st.progress(0)
            success_count = 0
            
            for i, f in enumerate(files):
                df = extract_data(f.name, f.read(), pwd_input, cols_map)
                if df is not None:
                    df['Source'] = f.name
                    all_dfs.append(df)
                    success_count += 1
                else:
                    st.error(f"檔案 {f.name} 讀取失敗，找不到欄位: {input_school}")
                bar.progress((i+1)/len(files))
            
            if all_dfs:
                st.session_state.merged_df = pd.concat(all_dfs, ignore_index=True)
                # 填充空值
                st.session_state.merged_df.fillna("未知", inplace=True)
                st.success(f"成功讀取 {success_count} 個檔案！")
                st.session_state.step = 2
                st.rerun()
            else:
                st.error("所有檔案都讀取失敗，請檢查左側欄位名稱設定是否與 Excel 表頭完全一致。")
        else:
            st.warning("請先上傳檔案")

# --- 步驟 2: 設定統計邏輯 ---
if st.session_state.step == 2:
    st.header("步驟 2: 定義身分 (誰是老師？誰是學生？)")
    
    df = st.session_state.merged_df
    
    # 抓出 Excel 裡所有出現過的職稱 (轉字串避免報錯)
    unique_roles = [str(r) for r in df['Role'].unique().tolist() if str(r).strip() != '']
    
    if not unique_roles:
        st.warning("⚠️ 警告：職稱欄位似乎是空的，無法區分師生。所有人都將計入「總參與人數」。")
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("🧑‍🏫 哪些職稱算「老師/成人」？")
        # 修正：加入 unique key
        teacher_list = st.multiselect(
            "請勾選老師職稱 (可多選)", 
            options=unique_roles,
            default=[r for r in unique_roles if '師' in str(r) or '長' in str(r) or '教' in str(r)],
            key="ms_teachers" 
        )
        
    with col2:
        st.subheader("👶 哪些職稱算「學生」？")
        # 修正：加入 unique key
        student_list = st.multiselect(
            "請勾選學生職稱 (可多選)", 
            options=unique_roles,
            default=[r for r in unique_roles if '生' in str(r)],
            key="ms_students"
        )

    st.info(f"目前資料共 {len(df)} 筆。")

    if st.button("開始計算統計報表 🚀", type="primary"):
        # 開始進行聚合分析
        
        # 1. 標記
        def tag_role(role_val):
            r = str(role_val)
            if r in teacher_list: return 'Teacher'
            if r in student_list: return 'Student'
            return 'Other'
            
        df['Tag'] = df['Role'].apply(tag_role)
        
        # 2. 統計總表
        summary_list = []
        
        for city, group in df.groupby('City'):
            school_count = group['School'].nunique()
            teacher_count = len(group[group['Tag'] == 'Teacher'])
            student_count = len(group[group['Tag'] == 'Student'])
            total_count = len(group) # 包含未勾選的人
            
            summary_list.append({
                '縣市': city,
                '學校數': school_count,
                '老師人數': teacher_count,
                '學生人數': student_count,
                '總參與人數': total_count
            })
            
        summary_df = pd.DataFrame(summary_list)
        if not summary_df.empty:
            summary_df = summary_df.sort_values(by='總參與人數', ascending=False).reset_index(drop=True)
        st.session_state.summary_df = summary_df
        
        # 3. 學校明細表
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
    
    if summary_df.empty:
        st.error("統計結果為空，請檢查您的職稱勾選設定。")
    else:
        # 1. 顯示總表
        st.subheader("🏆 縣市統計總表")
        st.dataframe(summary_df, use_container_width=True)
        
        # 2. 圖表
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("##### 各縣市學校數")
            st.bar_chart(summary_df.set_index('縣市')['學校數'])
        with c2:
            st.markdown("##### 各縣市參與人數")
            st.bar_chart(summary_df.set_index('縣市')['總參與人數'])

        # 3. 下載
        st.divider()
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            summary_df.to_excel(writer, sheet_name='縣市統計總表', index=False)
            school_df.to_excel(writer, sheet_name='各校詳細統計', index=False)
            st.session_state.merged_df.to_excel(writer, sheet_name='原始合併名單', index=False)
            
            # 美化
            workbook = writer.book
            fmt_header = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
            
            for sheet_name in ['縣市統計總表', '各校詳細統計']:
                ws = writer.sheets[sheet_name]
                ws.set_column(0, 5, 15)
            
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
