import streamlit as st
import os
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

# Информация о навигации
st.info("💡 **Навигация:** Нажмите на кнопку под карточкой инструмента для перехода.")
st.markdown("")

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
                                
            # Кнопка-ссылка для открытия в новом окне
            page_name = tool['page']
            
            # Определяем URL в зависимости от окружения
            if os.getenv('STREAMLIT_SERVER_BASE_URL') or os.getenv('STREAMLIT_SHARING'):
                # На Streamlit Cloud
                base_url = "https://client-analytics.streamlit.app"
            else:
                # Локально - используем относительный путь
                base_url = ""
            
            page_url = f"{base_url}/pages/{page_name}" if base_url else f"/pages/{page_name}"
            
            # Создаем стилизованную кнопку-ссылку, которая откроется в новом окне
            st.markdown(f"""
            <div style="text-align: center; margin-top: 15px;">
                <a href="{page_url}" target="_blank" rel="noopener noreferrer" style="
                    display: inline-block;
                    width: 100%;
                    padding: 12px 30px;
                    background-color: #4CAF50;
                    color: white !important;
                    text-decoration: none;
                    border-radius: 8px;
                    font-weight: bold;
                    text-align: center;
                    transition: background-color 0.3s ease;
                    cursor: pointer;
                    box-shadow: 0 2px 4px rgba(0,0,0,0.2);
                " onmouseover="this.style.backgroundColor='#45a049'; this.style.boxShadow='0 4px 8px rgba(0,0,0,0.3)'" 
                   onmouseout="this.style.backgroundColor='#4CAF50'; this.style.boxShadow='0 2px 4px rgba(0,0,0,0.2)'">
                    🔄 Открыть инструмент (в новом окне)
                </a>
            </div>
            """, unsafe_allow_html=True)
            
            # Альтернативная ссылка для открытия в текущем окне
            st.markdown(f"""
            <div style="text-align: center; margin-top: 10px;">
                <a href="{page_url}" target="_self" style="
                    color: #4CAF50;
                    text-decoration: none;
                    font-size: 0.9em;
                ">Или откройте в текущем окне</a>
            </div>
            """, unsafe_allow_html=True)

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
