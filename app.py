import streamlit as st
from config import PAGE_CONFIG

# Настройка страницы
st.set_page_config(**PAGE_CONFIG)

# CSS стили для главной страницы
st.markdown("""
<style>
    .main-header {
        text-align: center;
        padding: 20px 0;
        margin-bottom: 30px;
    }
    .tool-card {
        border: 2px solid #e0e0e0;
        border-radius: 15px;
        padding: 25px;
        margin: 15px 0;
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        text-align: center;
        transition: transform 0.3s ease, box-shadow 0.3s ease;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    .tool-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 8px 15px rgba(0, 0, 0, 0.2);
    }
    .tool-icon {
        font-size: 3em;
        margin: 15px 0;
    }
    .tool-name {
        font-size: 1.3em;
        font-weight: bold;
        margin: 15px 0;
        color: #2c3e50;
    }
    .tool-description {
        color: #555;
        margin: 10px 0;
        font-size: 0.95em;
    }
    .stButton > button {
        width: 100%;
        background-color: #4CAF50;
        color: white;
        font-weight: bold;
        padding: 10px 20px;
        border-radius: 8px;
        border: none;
        transition: background-color 0.3s ease;
    }
    .stButton > button:hover {
        background-color: #45a049;
    }
</style>
""", unsafe_allow_html=True)

# Главная страница
st.markdown('<div class="main-header">', unsafe_allow_html=True)
st.title("📊 Клиентская аналитика")
st.markdown('</div>', unsafe_allow_html=True)

st.markdown("---")

# Описание
st.markdown("""
<div style="text-align: center; font-size: 1.1em; color: #555; margin-bottom: 30px;">
    Добро пожаловать в систему клиентской аналитики!<br>
    Выберите инструмент для работы:
</div>
""", unsafe_allow_html=True)

# Список доступных инструментов
tools = [
    {
        "name": "Когортный анализ, возвращаемость и отток",
        "icon": "📊",
        "description": "Анализ когорт клиентов, возвращаемость и отток",
        "page": "cohort_analysis"
    }
]

# Создаем кнопки для каждого инструмента
st.markdown("### 🛠️ Доступные инструменты")
st.markdown("")

# Используем колонки для красивого размещения кнопок
for i in range(0, len(tools), 2):
    cols = st.columns(2)
    for j, tool in enumerate(tools[i:i+2]):
        with cols[j]:
            # Создаем карточку инструмента
            st.markdown(f"""
            <div class="tool-card">
                <div class="tool-icon">{tool['icon']}</div>
                <div class="tool-name">{tool['name']}</div>
                <div class="tool-description">{tool['description']}</div>
            </div>
            """, unsafe_allow_html=True)
            
            # Кнопка для перехода к инструменту
            # В Streamlit Pages правильный формат: "pages/filename" (без расширения .py)
            if st.button(f"Открыть инструмент", key=f"btn_{i+j}", use_container_width=True):
                # Используем правильный формат для Streamlit Pages
                # Формат должен быть точно: "pages/cohort_analysis" (без .py)
                st.switch_page(f"pages/{tool['page']}")

# Если инструментов нечетное количество, добавляем пустую колонку
if len(tools) % 2 == 1:
    st.markdown("")

# Информация о системе
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #666; padding: 20px;">
    <p>Система клиентской аналитики</p>
</div>
""", unsafe_allow_html=True)
