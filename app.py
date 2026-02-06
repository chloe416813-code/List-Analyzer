import streamlit as st
import pandas as pd
import io
import xlsxwriter

# ================= 0. 系統設定 =================
st.set_page_config(page_title="科普列車 - 純淨統計系統", page_icon="📊", layout="wide")

# 初始化 Session State
if 'df_result' not in st.session_state:
    st.session_state.df_result = pd.DataFrame()
if 'analysis_done' not in st.session_state:
    st.session_state.analysis_done = False

# ================= 1. 核心邏輯區 =================

def find_column_name(columns, keywords):
    """
    智慧欄位搜尋：
    在 Excel 的所有欄位名稱中，尋找包含指定關鍵字的欄位。
    """
    # 1. 先找完全符合的
    for col in columns:
        if str(col).strip() in keywords:
            return col
    # 2. 再找模糊符合的 (包含關鍵字)
    for col in columns:
        for k in keywords:
            if k in str(col).strip():
                return col
    return None

def process_files(files):
    """
    讀取多個檔案，只保留 [縣市, 學校, 職稱] 三個欄位，其餘丟棄
    """
    all_data = []
    
    # 定義我們要抓的關鍵字
    keys_city = ['縣市', '城市', 'City', '地區', '居住地', '地址']
    keys_school = ['學校', '校名', 'School', '單位', '就讀', '服務單位']
    keys_role = ['職稱', '身分', '身份', 'Role', '職務', '對象', '類別']
    
    for file in files:
        try:
            # 讀取 Excel
            df = pd.read_excel(file)
            
            # 搜尋欄位
            col_city = find_column_name(df.columns, keys_city)
            col_school = find_column_name(df.columns, keys_school)
            col_role = find_column_name(df.columns, keys_role)
            
            # 建立乾淨的資料表 (只取這三欄)
            clean_df = pd.DataFrame()
            
            if col_city: clean_df['縣市'] = df[col_city]
            else: clean_df['縣市'] = '未填縣市'
                
            if col_school: clean_df['學校'] = df[col_school]
            else: clean_df['學校'] = '未填學校'
                
            if col_role: clean_df['職稱'] = df[col_role]
            else: clean_df['職稱'] = '一般人員'
            
            # 標記來源 (方便除錯)
            clean_df['來源檔案'] = file.name
            
            all_data.append(clean_df)
            
        except Exception as e:
            st.error(f"檔案 {file.name} 讀取失敗: {e}")
            
    if all_data:
        return pd.concat(all_data, ignore_index=True)
    else:
        return pd.DataFrame()

# ================= 2. 介面設計 =================

st.title("📊 科普列車 - 快速人數統計系統")
st.markdown("### 自動抓取：縣市、學校、職稱")
st.info("此程式會自動忽略 Excel 中的其他個資欄位，只針對人數進行分析。")

# 上傳區
files = st.file_uploader("請上傳 Excel 名單 (可多選合併)", type=['xlsx'], accept_multiple_files=True)

if st.button("🚀 開始分析", type="primary"):
    if not files:
        st.warning("請先上傳檔案！")
    else:
        with st.spinner("正在清洗資料與計算人數..."):
            # 1. 執行處理
            result_df = process_files(files)
            
            if not result_df.empty:
                # 填充空值
                result_df.fillna("未知", inplace=True)
                st.session_state.df_result = result_df
                st.session_state.analysis_done = True
            else:
                st.error("無法抓取有效資料，請檢查 Excel 表頭名稱。")

# ================= 3. 分析結果儀表板 =================

if st.session_state.analysis_done and not st.session_state.df_result.empty:
    df = st.session_state.df_result
    
    st.divider()
    
    # --- 頂部數據卡 ---
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("總人數", f"{len(df)} 人")
    col2.metric("涵蓋縣市", f"{df['縣市'].nunique()} 個")
    col3.metric("參與學校", f"{df['學校'].nunique()} 所")
    col4.metric("檔案數量", f"{df['來源檔案'].nunique()} 個")
    
    # --- 圖表區 ---
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("🌍 各縣市人數分佈")
        # 畫長條圖
        city_counts = df['縣市'].value_counts()
        st.bar_chart(city_counts)
        
    with c2:
        st.subheader("🎓 師生/職稱比例")
        role_counts = df['職稱'].value_counts()
        st.bar_chart(role_counts, color="#ffaa00")

    # --- 詳細統計表 (樞紐分析) ---
    st.divider()
    st.subheader("🏫 各校詳細統計表")
    
    try:
        # 製作樞紐分析表：列=[縣市, 學校], 欄=[職稱], 值=人數
        pivot_df = df.pivot_table(index=['縣市', '學校'], columns='職稱', aggfunc='size', fill_value=0)
        
        # 加入總計欄
        pivot_df['該校總計'] = pivot_df.sum(axis=1)
        
        # 顯示在網頁上
        st.dataframe(pivot_df, use_container_width=True)
        
        # --- 下載區 ---
        st.divider()
        st.subheader("📥 下載報表")
        
        # 製作 Excel 下載檔
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Sheet 1: 樞紐分析 (最重要)
            pivot_df.to_excel(writer, sheet_name='各校師生統計')
            
            # Sheet 2: 簡單的縣市統計
            df['縣市'].value_counts().to_frame(name="人數").to_excel(writer, sheet_name='縣市統計')
            
            # Sheet 3: 職稱統計
            df['職稱'].value_counts().to_frame(name="人數").to_excel(writer, sheet_name='職稱統計')
            
            # Sheet 4: 乾淨的總名單 (不含個資)
            df.to_excel(writer, sheet_name='總名單明細', index=False)
            
            # 美化欄寬
            writer.sheets['各校師生統計'].set_column(0, 1, 20) # 縣市, 學校欄寬

        st.download_button(
            label="📊 下載 Excel 統計報告",
            data=output.getvalue(),
            file_name="科普列車_統計分析.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
    except Exception as e:
        st.error(f"統計表產生錯誤: {e}")
