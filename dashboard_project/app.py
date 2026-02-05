import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# --- НАСТРОЙКИ СТРАНИЦЫ ---
st.set_page_config(page_title="Executive Sales Dashboard", layout="wide")

# --- ФУНКЦИИ ЗАГРУЗКИ И ОБРАБОТКИ ---

def find_header_row(df, keywords):
    """
    Ищет строку, в которой содержится большинство ключевых слов.
    Это позволяет игнорировать строки типа 'Last updated' или пустые строки сверху.
    """
    for i in range(min(20, len(df))):  # Проверяем первые 20 строк
        row_values = df.iloc[i].astype(str).tolist()
        # Считаем совпадения ключевых слов в строке
        matches = sum(1 for k in keywords if any(k.lower() in val.lower() for val in row_values))
        if matches >= 2:  # Если нашли хотя бы 2 ключевых слова в строке
            return i
    return 0

def clean_currency(x):
    """Очистка данных от знаков валют и пробелов"""
    if isinstance(x, str):
        clean_str = x.replace('$', '').replace('€', '').replace(',', '').replace(' ', '')
        try:
            return float(clean_str)
        except ValueError:
            return 0.0
    return x

@st.cache_data
def load_data(uploaded_file):
    try:
        # Читаем Excel файл (все листы сразу)
        xls = pd.ExcelFile(uploaded_file)
        
        # 1. Пытаемся найти лист со сводными данными (Consolidated)
        # Ищем лист, в названии которого есть 'Consolidated' или 'Total'
        target_sheet = next((s for s in xls.sheet_names if 'Consolidated' in s), None)
        
        if not target_sheet:
            target_sheet = xls.sheet_names[0] # Если не нашли, берем первый
            
        df_raw = pd.read_excel(uploaded_file, sheet_name=target_sheet, header=None)
        
        # Умный поиск заголовка. Ключевые слова из твоих файлов.
        keywords = ['Sales Manager', 'Region', 'Brand', 'Sales 2024', 'Forecast']
        header_idx = find_header_row(df_raw, keywords)
        
        # Перезагружаем с правильным заголовком
        df = pd.read_excel(uploaded_file, sheet_name=target_sheet, header=header_idx)
        
        # Очистка названий колонок (убираем пробелы и переносы строк)
        df.columns = df.columns.astype(str).str.strip().str.replace('\n', ' ')
        
        # Стандартизация важных колонок (ищем похожие названия)
        col_map = {}
        for col in df.columns:
            if 'Region' in col: col_map[col] = 'Region'
            elif 'Brand' in col: col_map[col] = 'Brand'
            elif 'Manager' in col: col_map[col] = 'Manager'
            elif 'Forecast 2026' in col: col_map[col] = 'Forecast'
            elif 'Target 2026' in col: col_map[col] = 'Target'
            elif 'Sales 2025' in col: col_map[col] = 'Sales_Prev'
        
        df = df.rename(columns=col_map)
        
        # Фильтрация "мусорных" строк (итогов и пустых)
        if 'Region' in df.columns:
            df = df[df['Region'].notna()]
            df = df[~df['Region'].astype(str).str.contains('Total', case=False)]
            
        # Преобразование чисел
        numeric_cols = ['Forecast', 'Target', 'Sales_Prev']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = df[col].apply(clean_currency).fillna(0)
                
        return df
        
    except Exception as e:
        st.error(f"Ошибка при чтении файла: {e}")
        return None

# --- ИНТЕРФЕЙС ---

st.title("📊 Корпоративный Дашборд Продаж 2026")
st.markdown("Инструмент для анализа выполнения плана и моделирования сценариев.")

# --- САЙДБАР ---
with st.sidebar:
    st.header("1. Загрузка данных")
    uploaded_file = st.file_uploader("Загрузите Excel файл (Sales Dashboard 2026)", type=["xlsx", "xls"])
    
    st.header("2. Моделирование сценариев")
    st.info("Измените параметры, чтобы увидеть влияние на прогноз:")
    
    price_impact = st.slider("Изменение цен (%)", -20, 20, 0, 1)
    traffic_impact = st.slider("Рост объема заказов (%)", -20, 50, 0, 1)
    conversion_rate = st.slider("Коэф. успешных сделок", 0.5, 1.5, 1.0, 0.1)

# --- ОСНОВНАЯ ЛОГИКА ---
if uploaded_file is not None:
    df = load_data(uploaded_file)
    
    if df is not None:
        # --- ФИЛЬТРЫ ---
        with st.expander("🔎 Фильтры данных", expanded=True):
            col1, col2, col3 = st.columns(3)
            
            selected_region = "Все"
            selected_brand = "Все"
            selected_manager = "Все"
            
            if 'Region' in df.columns:
                regions = ["Все"] + sorted(df['Region'].unique().astype(str).tolist())
                selected_region = col1.selectbox("Регион", regions)
            
            if 'Brand' in df.columns:
                brands = ["Все"] + sorted(df['Brand'].unique().astype(str).tolist())
                selected_brand = col2.selectbox("Бренд", brands)
                
            if 'Manager' in df.columns:
                managers = ["Все"] + sorted(df['Manager'].unique().astype(str).tolist())
                selected_manager = col3.selectbox("Менеджер", managers)

        # Применение фильтров
        df_filtered = df.copy()
        if selected_region != "Все":
            df_filtered = df_filtered[df_filtered['Region'] == selected_region]
        if selected_brand != "Все":
            df_filtered = df_filtered[df_filtered['Brand'] == selected_brand]
        if selected_manager != "Все":
            df_filtered = df_filtered[df_filtered['Manager'] == selected_manager]

        # --- РАСЧЕТ KPI С УЧЕТОМ СЦЕНАРИЕВ ---
        # Логика модели: (Базовый прогноз * (1 + Цены) * (1 + Трафик)) * Конверсия
        
        # Коэффициенты
        p_factor = 1 + (price_impact / 100)
        t_factor = 1 + (traffic_impact / 100)
        
        # Проверка наличия колонок
        has_forecast = 'Forecast' in df_filtered.columns
        has_target = 'Target' in df_filtered.columns
        
        total_forecast_raw = df_filtered['Forecast'].sum() if has_forecast else 0
        total_target = df_filtered['Target'].sum() if has_target else 0
        
        # Моделируемый прогноз
        modeled_forecast = total_forecast_raw * p_factor * t_factor * conversion_rate
        
        delta_val = modeled_forecast - total_target
        
        # --- ОТОБРАЖЕНИЕ KPI ---
        st.divider()
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        
        kpi1.metric("Цель (Target 2026)", f"€ {total_target:,.0f}")
        kpi2.metric("Текущий Прогноз (Факт)", f"€ {total_forecast_raw:,.0f}")
        kpi3.metric("Моделируемый Итог", f"€ {modeled_forecast:,.0f}", 
                    delta=f"{((modeled_forecast/total_forecast_raw)-1)*100:.1f}% от факта" if total_forecast_raw else None)
        kpi4.metric("Отклонение от Плана", f"€ {delta_val:,.0f}", 
                    delta_color="normal" if delta_val >= 0 else "inverse")

        st.divider()

        # --- ГРАФИКИ (TABS) ---
        tab1, tab2, tab3 = st.tabs(["📈 Анализ Структуры", "🏆 Рейтинги", "📄 Данные"])
        
        with tab1:
            col_chart1, col_chart2 = st.columns(2)
            
            if 'Brand' in df_filtered.columns and has_forecast:
                fig_pie = px.pie(df_filtered, values='Forecast', names='Brand', 
                                title='Доля продаж по Брендам', hole=0.4)
                col_chart1.plotly_chart(fig_pie, use_container_width=True)
                
            if 'Region' in df_filtered.columns and has_forecast:
                fig_bar = px.bar(df_filtered.groupby('Region')['Forecast'].sum().reset_index(), 
                                x='Region', y='Forecast', 
                                title='Прогноз продаж по Регионам', color='Region')
                col_chart2.plotly_chart(fig_bar, use_container_width=True)

        with tab2:
            if 'Manager' in df_filtered.columns and has_forecast:
                manager_perf = df_filtered.groupby('Manager')[['Forecast', 'Target']].sum().reset_index()
                manager_perf['Achievement %'] = (manager_perf['Forecast'] / manager_perf['Target']) * 100
                manager_perf = manager_perf.sort_values('Forecast', ascending=True)
                
                fig_manager = go.Figure()
                fig_manager.add_trace(go.Bar(y=manager_perf['Manager'], x=manager_perf['Forecast'], 
                                            name='Прогноз', orientation='h'))
                fig_manager.add_trace(go.Bar(y=manager_perf['Manager'], x=manager_perf['Target'], 
                                            name='План', orientation='h'))
                
                fig_manager.update_layout(title="Эффективность Менеджеров (План vs Факт)", barmode='group')
                st.plotly_chart(fig_manager, use_container_width=True)
            else:
                st.warning("Недостаточно данных для построения рейтинга менеджеров")

        with tab3:
            st.dataframe(df_filtered, use_container_width=True)
            
    else:
        st.warning("Пожалуйста, загрузите файл. Убедитесь, что в файле есть лист 'Consolidated' или сводные данные.")
else:

    st.info("⬅️ Загрузите файл Excel в меню слева для начала работы.")
