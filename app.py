import streamlit as st

# Настройка страницы
st.set_page_config(
    page_title="Аналитика клиентов",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Аналитика клиентов")
st.markdown("---")

st.markdown("""
### Выберите инструмент для анализа:
""")

col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("""
    <a href="?page=pages/1__Когортный_анализ" style="text-decoration: none; color: inherit; display: block;">
        <div style="text-align: center; padding: 20px; border: 2px solid #4CAF50; border-radius: 10px; background: #f1f8f4; cursor: pointer; transition: all 0.3s ease;" 
             onmouseover="this.style.transform='translateY(-5px)'; this.style.boxShadow='0 8px 16px rgba(76, 175, 80, 0.2)';" 
             onmouseout="this.style.transform='translateY(0)'; this.style.boxShadow='none';">
            <h2>📊 Когортный анализ</h2>
            <p>Анализ возвращаемости и оттока клиентов</p>
            <p style="color: #4CAF50; font-weight: bold; margin-top: 10px;">✓ Доступен</p>
        </div>
    </a>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("""
    <div style="text-align: center; padding: 20px; border: 2px solid #e0e0e0; border-radius: 10px; background: #f8f9fa;">
        <h2>🔄 Цикл жизни клиента</h2>
        <p>Анализ жизненного цикла клиента на продукте</p>
        <p style="color: #999; font-size: 0.9em; margin-top: 10px;">🚧 В разработке</p>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown("""
    <div style="text-align: center; padding: 20px; border: 2px solid #e0e0e0; border-radius: 10px; background: #f8f9fa;">
        <h2>📈 Другие инструменты</h2>
        <p>Дополнительные аналитические инструменты</p>
        <p style="color: #999; font-size: 0.9em; margin-top: 10px;">🚧 В разработке</p>
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")
st.info("💡 Нажмите на карточку инструмента для перехода к нему")
