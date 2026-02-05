import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# --- НАСТРОЙКИ СТРАНИЦЫ ---
st.set_page_config(page_title="Executive Sales Dashboard", layout="wide")

# --- ФУНКЦИИ ЗАГРУЗКИ И ОБРАБОТКИ ---

def find_header_row(df, keywords):
    """Ищет строку заголовка по ключевым словам"""
    for i in range(min(20, len(df))):
        row_values = df.iloc[i].astype(str).tolist()
        matches = sum(1 for k in keywords if any(k.lower() in val.lower() for val in row_values))
        if matches >= 2:
            return i
    return 0

def clean_currency(x):
    """Очистка от валют и пробелов перед конвертацией"""
    if isinstance(x, str):
        # Удаляем всё лишнее, оставляем цифры и точку
        clean_str = x.replace('$', '').replace('€', '').replace(',', '').replace(' ', '').strip()
        # Если пусто или '-' (часто в отчетах), возвращаем 0
        if not clean_str or clean_str == '-':
            return 0
        return clean_str
    return x

@st.cache_data
def load_data(uploaded_file):
    try:
        xls = pd.ExcelFile(uploaded_file)
        
        # 1. Ищем лист Consolidated
        target_sheet = next((s for s in xls.sheet_names if 'Consolidated' in s), None)
        if not target_sheet:
            target_sheet = xls.sheet_names[0]
            
        df_raw = pd.read_excel(uploaded_file, sheet_name=target_sheet, header=None)
        
        # 2. Умный поиск заголовка
        keywords = ['Sales Manager', 'Region', 'Brand', 'Sales 2024', 'Forecast']
        header_idx = find_header_row(df_raw, keywords)
        
        # 3. Читаем данные правильно
        df = pd.read_excel(uploaded_file, sheet_name=target_sheet, header=header_idx)
        
        # 4. Очистка имен колонок
        df.columns = df.columns.astype(str).str.strip().str.replace('\n', ' ')
        
        # 5. Карта переименования (для стандартизации)
        col_map = {}
        for col in df.columns:
            if 'Region' in col: col_map[col] = 'Region'
            elif 'Brand' in col: col_map[col] = 'Brand'
            elif 'Manager' in col: col_map[col] = 'Manager'
            elif 'Forecast 2026' in col: col_map[col] = 'Forecast' # Важно: точное совпадение
            elif 'Forecast' in col and 'Target' not in col: col_map[col] = 'Forecast' # Если имя другое
            elif 'Target 2026' in col: col_map[col] = 'Target'
            elif 'Sales 2025' in col: col_map[col] = 'Sales_Prev'
        
        df = df.rename(columns=col_map)
        
        # 6. Фильтрация итоговых строк
        if 'Region' in df.columns:
            df = df[df['Region'].notna()]
            df = df[~df['Region'].astype(str).str.contains('Total', case=False, na=False)]
            
        # 7. ЖЕСТКОЕ ПРЕОБРАЗОВАНИЕ ЧИСЕЛ (Fix TypeError)
        numeric_cols = ['Forecast', 'Target', 'Sales_Prev']
        for col in numeric_cols:
            if col in df.columns:
                # Сначала чистим символы
                df[col] = df[col].apply(clean_currency)
                # Затем принудительно в числа (ошибки -> NaN -> 0)
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                
        return df
        
    except Exception as e:
        st.error(f"Ошибка обработки файла: {e}")
        return None

# --- ИНТЕРФЕЙС ---

st.title("📊 Корпоративный Дашборд Продаж 2026")

with st.sidebar:
    st.header("1. Загрузка данных")
    uploaded_file = st.file_uploader("Загрузите Excel файл", type=["xlsx", "xls"])
    
    st.header("2. Сценарии")
    price_impact = st.slider("Изменение цен (%)", -20, 20, 0, 1)
    traffic_impact = st.slider("Рост объема (%)", -20, 50, 0, 1)
    conversion_rate = st.slider("Конверсия", 0.5, 1.5, 1.0, 0.1)

if uploaded_file is not None:
    df = load_data(uploaded_file)
    
    if df is not None:
        # Фильтры
        with st.expander("🔎 Фильтры", expanded=True):
            c1, c2, c3 = st.columns(3)
            # Безопасное получение списков (с проверкой на наличие колонки)
            regions = ["Все"] + sorted(df['Region'].unique().astype(str).tolist()) if 'Region' in df else ["Все"]
            brands = ["Все"] + sorted(df['Brand'].unique().astype(str).tolist()) if 'Brand' in df else ["Все"]
            managers = ["Все"] + sorted(df['Manager'].unique().astype(str).tolist()) if 'Manager' in df else ["Все"]
            
            sel_region = c1.selectbox("Регион", regions)
            sel_brand = c2.selectbox("Бренд", brands)
            sel_manager = c3.selectbox("Менеджер", managers)

        # Применение фильтров
        df_filtered = df.copy()
        if 'Region' in df and sel_region != "Все":
            df_filtered = df_filtered[df_filtered['Region'] == sel_region]
        if 'Brand' in df and sel_brand != "Все":
            df_filtered = df_filtered[df_filtered['Brand'] == sel_brand]
        if 'Manager' in df and sel_manager != "Все":
            df_filtered = df_filtered[df_filtered['Manager'] == sel_manager]

        # Расчеты
        has_forecast = 'Forecast' in df_filtered.columns
        has_target = 'Target' in df_filtered.columns
        
        # Безопасное суммирование (теперь данные точно числа)
        total_forecast = df_filtered['Forecast'].sum() if has_forecast else 0.0
        total_target = df_filtered['Target'].sum() if has_target else 0.0
        
        # Модель
        modeled = total_forecast * (1 + price_impact/100) * (1 + traffic_impact/100) * conversion_rate
        delta = modeled - total_target
        
        # KPI
        st.divider()
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("План (Target)", f"€ {total_target:,.0f}")
        k2.metric("Прогноз (Forecast)", f"€ {total_forecast:,.0f}")
        k3.metric("Модель", f"€ {modeled:,.0f}", 
                  delta=f"{((modeled/total_forecast)-1)*100:.1f}%" if total_forecast else None)
        k4.metric("Отклонение", f"€ {delta:,.0f}", delta_color="normal" if delta >= 0 else "inverse")
        st.divider()

        # Графики
        tab1, tab2, tab3 = st.tabs(["Динамика", "Рейтинг", "Данные"])
        
        with tab1:
            col1, col2 = st.columns(2)
            if 'Brand' in df_filtered and has_forecast:
                fig = px.pie(df_filtered, values='Forecast', names='Brand', title='Доля по Брендам', hole=0.4)
                col1.plotly_chart(fig, use_container_width=True)
            if 'Region' in df_filtered and has_forecast:
                fig = px.bar(df_filtered.groupby('Region')['Forecast'].sum().reset_index(), 
                             x='Region', y='Forecast', title='По Регионам')
                col2.plotly_chart(fig, use_container_width=True)
                
        with tab2:
            if 'Manager' in df_filtered and has_forecast and has_target:
                m_df = df_filtered.groupby('Manager')[['Forecast', 'Target']].sum().reset_index()
                m_df = m_df.sort_values('Forecast', ascending=True)
                
                fig = go.Figure()
                fig.add_trace(go.Bar(y=m_df['Manager'], x=m_df['Forecast'], name='Прогноз', orientation='h'))
                fig.add_trace(go.Bar(y=m_df['Manager'], x=m_df['Target'], name='План', orientation='h'))
                fig.update_layout(title="Эффективность Менеджеров")
                st.plotly_chart(fig, use_container_width=True)
                
        with tab3:
            st.dataframe(df_filtered, use_container_width=True)

    else:
        st.warning("Файл загружен, но данные не распознаны. Проверьте названия колонок (Consolidated).")
else:
    st.info("⬅️ Загрузите файл Excel")
