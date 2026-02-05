import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# --- НАСТРОЙКИ СТРАНИЦЫ ---
st.set_page_config(page_title="Executive Sales Dashboard", layout="wide")

# --- ФУНКЦИИ ---

def find_header_row(df, keywords):
    """Ищет индекс строки заголовка"""
    for i in range(min(20, len(df))):
        row_values = df.iloc[i].astype(str).tolist()
        matches = sum(1 for k in keywords if any(k.lower() in val.lower() for val in row_values))
        if matches >= 2:
            return i
    return 0

def clean_currency(x):
    """Очистка строк от валют и пробелов"""
    if isinstance(x, str):
        clean = x.replace('$', '').replace('€', '').replace(',', '').replace(' ', '').strip()
        if not clean or clean in ['-', 'nan', 'None']:
            return 0
        return clean
    return x

@st.cache_data
def load_data(uploaded_file):
    try:
        xls = pd.ExcelFile(uploaded_file)
        
        # 1. Поиск листа Consolidated
        target_sheet = next((s for s in xls.sheet_names if 'Consolidated' in s), None)
        if not target_sheet:
            target_sheet = xls.sheet_names[0]
            
        # 2. Поиск заголовка
        df_raw = pd.read_excel(uploaded_file, sheet_name=target_sheet, header=None)
        keywords = ['Sales Manager', 'Region', 'Brand', 'Sales 2024', 'Forecast']
        header_idx = find_header_row(df_raw, keywords)
        
        # 3. Чтение данных
        df = pd.read_excel(uploaded_file, sheet_name=target_sheet, header=header_idx)
        
        # 4. Очистка имен колонок
        df.columns = df.columns.astype(str).str.strip().str.replace('\n', ' ')
        
        # 5. УМНОЕ ПЕРЕИМЕНОВАНИЕ (С защитой от дубликатов)
        col_map = {}
        used_targets = set()
        
        for col in df.columns:
            new_name = None
            col_lower = col.lower()
            
            if 'region' in col_lower: new_name = 'Region'
            elif 'brand' in col_lower: new_name = 'Brand'
            elif 'manager' in col_lower: new_name = 'Manager'
            elif 'sales 2025' in col_lower: new_name = 'Sales_Prev'
            elif 'target 2026' in col_lower: new_name = 'Target'
            elif 'forecast 2026' in col_lower: new_name = 'Forecast'
            elif 'forecast' in col_lower and 'target' not in col_lower: 
                new_name = 'Forecast'
            
            if new_name:
                if new_name in used_targets:
                    continue
                col_map[col] = new_name
                used_targets.add(new_name)
        
        df = df.rename(columns=col_map)
        df = df.loc[:, ~df.columns.duplicated()]
        
        # 6. Фильтрация мусора
        if 'Region' in df.columns:
            df = df[df['Region'].notna()]
            df = df[~df['Region'].astype(str).str.contains('Total', case=False, na=False)]
            
        # 7. Конвертация чисел
        numeric_cols = ['Forecast', 'Target', 'Sales_Prev']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = df[col].astype(str).apply(clean_currency)
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        # 8. ВАЖНО: Сброс индекса, чтобы убрать "дырки" в нумерации
        df = df.reset_index(drop=True)
                
        return df
        
    except Exception as e:
        st.error(f"Ошибка чтения файла: {e}")
        return None

# --- ИНТЕРФЕЙС ---

st.title("📊 Корпоративный Дашборд Продаж 2026")

with st.sidebar:
    st.header("1. Загрузка")
    uploaded_file = st.file_uploader("Файл Excel (Sales Dashboard)", type=["xlsx", "xls"])
    
    st.header("2. Сценарии")
    price_impact = st.slider("Цена (%)", -20, 20, 0, 1)
    traffic_impact = st.slider("Объем (%)", -20, 50, 0, 1)
    conversion = st.slider("Конверсия", 0.5, 1.5, 1.0, 0.1)

if uploaded_file:
    df = load_data(uploaded_file)
    
    if df is not None:
        # --- ФИЛЬТРЫ ---
        with st.expander("🔎 Фильтры", expanded=True):
            c1, c2, c3 = st.columns(3)
            
            regions = ["Все"] + sorted(df['Region'].unique().astype(str)) if 'Region' in df else ["Все"]
            brands = ["Все"] + sorted(df['Brand'].unique().astype(str)) if 'Brand' in df else ["Все"]
            managers = ["Все"] + sorted(df['Manager'].unique().astype(str)) if 'Manager' in df else ["Все"]
            
            sel_region = c1.selectbox("Регион", regions)
            sel_brand = c2.selectbox("Бренд", brands)
            sel_manager = c3.selectbox("Менеджер", managers)

        # --- ПОШАГОВАЯ ФИЛЬТРАЦИЯ (Исправлено) ---
        # Теперь мы фильтруем таблицу шаг за шагом, это безопаснее
        df_filtered = df.copy()
        
        if 'Region' in df_filtered.columns and sel_region != "Все":
            df_filtered = df_filtered[df_filtered['Region'] == sel_region]
            
        if 'Brand' in df_filtered.columns and sel_brand != "Все":
            df_filtered = df_filtered[df_filtered['Brand'] == sel_brand]
            
        if 'Manager' in df_filtered.columns and sel_manager != "Все":
            df_filtered = df_filtered[df_filtered['Manager'] == sel_manager]

        # --- KPI ---
        has_forecast = 'Forecast' in df_filtered
        has_target = 'Target' in df_filtered
        
        total_fc = df_filtered['Forecast'].sum() if has_forecast else 0
        total_tg = df_filtered['Target'].sum() if has_target else 0
        
        # Расчет модели
        modeled = total_fc * (1 + price_impact/100) * (1 + traffic_impact/100) * conversion
        delta = modeled - total_tg
        
        st.divider()
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("План (Target)", f"€ {total_tg:,.0f}")
        k2.metric("Факт (Forecast)", f"€ {total_fc:,.0f}")
        k3.metric("Модель", f"€ {modeled:,.0f}", 
                  delta=f"{((modeled/total_fc)-1)*100:.1f}%" if total_fc else None)
        k4.metric("Отклонение", f"€ {delta:,.0f}", 
                  delta_color="normal" if delta >= 0 else "inverse")
        st.divider()

        # --- ГРАФИКИ ---
        t1, t2, t3 = st.tabs(["Динамика", "Рейтинг", "Данные"])
        
        with t1:
            c_g1, c_g2 = st.columns(2)
            if has_forecast and 'Brand' in df_filtered:
                # Группируем, чтобы убрать дубликаты в графике
                pie_data = df_filtered.groupby('Brand')['Forecast'].sum().reset_index()
                fig = px.pie(pie_data, values='Forecast', names='Brand', title='Продажи по Брендам', hole=0.4)
                c_g1.plotly_chart(fig, use_container_width=True)
                
            if has_forecast and 'Region' in df_filtered:
                reg_data = df_filtered.groupby('Region')['Forecast'].sum().reset_index()
                fig = px.bar(reg_data, x='Region', y='Forecast', title='Продажи по Регионам')
                c_g2.plotly_chart(fig, use_container_width=True)
                
        with t2:
            if has_forecast and has_target and 'Manager' in df_filtered:
                m_data = df_filtered.groupby('Manager')[['Forecast', 'Target']].sum().reset_index()
                m_data = m_data.sort_values('Forecast')
                
                fig = go.Figure()
                fig.add_trace(go.Bar(y=m_data['Manager'], x=m_data['Forecast'], name='Прогноз', orientation='h'))
                fig.add_trace(go.Bar(y=m_data['Manager'], x=m_data['Target'], name='План', orientation='h'))
                fig.update_layout(title="Эффективность Менеджеров", barmode='group')
                st.plotly_chart(fig, use_container_width=True)
                
        with t3:
            st.dataframe(df_filtered, use_container_width=True)
            
    else:
        st.warning("Не удалось прочитать данные. Проверьте лист 'Consolidated'.")
else:
    st.info("⬅️ Загрузите файл Excel")
