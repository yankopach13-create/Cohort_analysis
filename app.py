import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from datetime import datetime
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')  # Используем неинтерактивный бэкенд
import seaborn as sns
from reportlab.lib.pagesizes import A4, letter
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import platform
import os
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import platform
import os

# Настройка страницы
st.set_page_config(
    page_title="Когортный анализ",
    page_icon="📊",
    layout="wide"
)

# Функция для создания кнопки копирования в буфер обмена
def create_copy_button(text, button_label, key):
    """Создает кнопку для копирования текста в буфер обмена"""
    import streamlit.components.v1 as components
    import json
    
    # Экранируем текст для безопасной вставки в JavaScript
    # Используем JSON для правильного экранирования
    text_json = json.dumps(text)
    
    html = f"""
    <div style="width: 100%; margin: 5px 0;">
        <button onclick="copyToClipboard_{key}()" style="
            width: 100%;
            padding: 8px 12px;
            background: #f8f9fa;
            color: #333;
            border: 2px solid #e0e0e0;
            border-bottom: 2px solid #e0e0e0;
            border-radius: 8px;
            cursor: pointer;
            font-weight: 400;
            font-size: 0.85rem;
            line-height: 1.3;
            text-align: center;
            min-height: auto;
            height: auto;
            display: flex;
            align-items: center;
            justify-content: center;
            white-space: normal;
            word-wrap: break-word;
            overflow-wrap: break-word;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
            transition: all 0.3s ease;
            margin: 0;
            box-sizing: border-box;
        " onmouseover="this.style.transform='translateY(-2px)'; this.style.boxShadow='0 4px 8px rgba(0, 0, 0, 0.1)'; this.style.background='#ffffff'; this.style.borderColor='#d0d0d0'; this.style.borderBottomColor='#d0d0d0';" onmouseout="this.style.transform='translateY(0)'; this.style.boxShadow='0 2px 4px rgba(0, 0, 0, 0.05)'; this.style.background='#f8f9fa'; this.style.borderColor='#e0e0e0'; this.style.borderBottomColor='#e0e0e0';" onmousedown="this.style.transform='translateY(0)'; this.style.boxShadow='0 2px 4px rgba(0, 0, 0, 0.05)';" onmouseup="this.style.transform='translateY(-2px)'; this.style.boxShadow='0 4px 8px rgba(0, 0, 0, 0.1)';">
            <div style="display: flex; align-items: center; justify-content: center; width: 100%;">
                <p style="margin: 0; padding: 0; font-size: 0.85rem; font-weight: 400; line-height: 1.3; word-wrap: break-word; overflow-wrap: break-word; white-space: normal;">{button_label}</p>
            </div>
        </button>
        <div id="copy_status_{key}" style="margin-top: 5px; color: rgb(0, 128, 0); font-size: 0.8rem; display: none; text-align: center;">✓ Скопировано!</div>
    </div>
    <script>
        const textToCopy_{key} = {text_json};
        
        function copyToClipboard_{key}() {{
            const text = textToCopy_{key};
            
            // Пробуем использовать современный API
            if (navigator.clipboard && navigator.clipboard.writeText) {{
                navigator.clipboard.writeText(text).then(function() {{
                    const status = document.getElementById('copy_status_{key}');
                    status.style.display = 'block';
                    setTimeout(function() {{
                        status.style.display = 'none';
                    }}, 2000);
                }}).catch(function(err) {{
                    // Fallback на старый метод
                    fallbackCopy_{key}(text);
                }});
            }} else {{
                // Fallback для старых браузеров
                fallbackCopy_{key}(text);
            }}
        }}
        
        function fallbackCopy_{key}(text) {{
            const textarea = document.createElement('textarea');
            textarea.value = text;
            textarea.style.position = 'fixed';
            textarea.style.left = '-999999px';
            textarea.style.top = '-999999px';
            document.body.appendChild(textarea);
            textarea.focus();
            textarea.select();
            try {{
                const successful = document.execCommand('copy');
                if (successful) {{
                    const status = document.getElementById('copy_status_{key}');
                    status.style.display = 'block';
                    setTimeout(function() {{
                        status.style.display = 'none';
                    }}, 2000);
                }} else {{
                    alert('Не удалось скопировать. Пожалуйста, скопируйте вручную.');
                }}
            }} catch(err) {{
                alert('Ошибка копирования: ' + err);
            }} finally {{
                document.body.removeChild(textarea);
            }}
        }}
    </script>
    """
    components.html(html, height=60)

st.title("📊 Когортный анализ, возвращаемость и отток")
st.markdown("---")

# Глобальные CSS стили для всех таблиц (выравнивание по центру)
st.markdown("""
<style>
div[data-testid="stDataFrame"] table,
div[data-testid="stDataFrame"] table th,
div[data-testid="stDataFrame"] table td {
    text-align: center !important;
}
div[data-testid="stDataFrame"] th,
div[data-testid="stDataFrame"] td {
    text-align: center !important;
}
</style>
""", unsafe_allow_html=True)

# Инициализация session state для хранения загруженных данных
if 'uploaded_data' not in st.session_state:
    st.session_state.uploaded_data = None
if 'df' not in st.session_state:
    st.session_state.df = None
if 'cohort_matrix' not in st.session_state:
    st.session_state.cohort_matrix = None
if 'cohort_info' not in st.session_state:
    st.session_state.cohort_info = None
if 'sorted_periods' not in st.session_state:
    st.session_state.sorted_periods = None
if 'year_month_col' not in st.session_state:
    st.session_state.year_month_col = None
if 'client_col' not in st.session_state:
    st.session_state.client_col = None

# Функция для преобразования периода (месяц или неделя) в порядковый номер для сортировки
def parse_period(period_str):
    """Преобразует период в кортеж для сортировки.
    Поддерживает форматы:
    - Месяцы: '2025-март', '2024-янв', '2024-январь'
    - Недели: '2025/01', '2024/52' (год/номер через слеш), '2024-W01', '2024-W1', '2024-нед01', '2024-нед1', '2024-н01'
    Возвращает (year, period_number, type) где type: 0=месяц, 1=неделя
    """
    try:
        period_str = str(period_str).strip()
        
        # Словарь месяцев
        months = {
            'янв': 1, 'январь': 1,
            'фев': 2, 'февраль': 2,
            'мар': 3, 'март': 3,
            'апр': 4, 'апрель': 4,
            'май': 5, 'май': 5,
            'июн': 6, 'июнь': 6,
            'июл': 7, 'июль': 7,
            'авг': 8, 'август': 8,
            'сен': 9, 'сентябрь': 9,
            'окт': 10, 'октябрь': 10,
            'ноя': 11, 'ноябрь': 11,
            'дек': 12, 'декабрь': 12
        }
        
        # Сначала пытаемся распарсить как месяц
        match_month = re.match(r'(\d{4})[-_]?([а-яА-Я]+)', period_str.lower())
        if match_month:
            year = int(match_month.group(1))
            month_name = match_month.group(2)
            month = months.get(month_name, 0)
            if month > 0:
                return (year, month, 0)  # 0 = месяц
        
        # Пытаемся распарсить как неделю в формате "2025/01" (год/номер недели через слеш)
        match_week_slash = re.match(r'(\d{4})[/](\d{1,2})$', period_str)
        if match_week_slash:
            year = int(match_week_slash.group(1))
            week = int(match_week_slash.group(2))
            if 1 <= week <= 53:
                return (year, week, 1)  # 1 = неделя
        
        # Пытаемся распарсить как неделю в формате ISO (2024-W01, 2024-W1)
        match_week_iso = re.match(r'(\d{4})[-_]?W(\d{1,2})', period_str.upper())
        if match_week_iso:
            year = int(match_week_iso.group(1))
            week = int(match_week_iso.group(2))
            if 1 <= week <= 53:
                return (year, week, 1)  # 1 = неделя
        
        # Пытаемся распарсить как неделю в формате "2024-нед01", "2024-нед1", "2024-н01"
        match_week_ru = re.match(r'(\d{4})[-_]?(?:нед|н)(\d{1,2})', period_str.lower())
        if match_week_ru:
            year = int(match_week_ru.group(1))
            week = int(match_week_ru.group(2))
            if 1 <= week <= 53:
                return (year, week, 1)  # 1 = неделя
        
        # Пытаемся распарсить как неделю в формате "2024-неделя01", "2024-неделя1"
        match_week_word = re.match(r'(\d{4})[-_]?неделя(\d{1,2})', period_str.lower())
        if match_week_word:
            year = int(match_week_word.group(1))
            week = int(match_week_word.group(2))
            if 1 <= week <= 53:
                return (year, week, 1)  # 1 = неделя
        
        # Пытаемся распарсить как "2024-01" - если число > 12, точно неделя, иначе нужно проверить контекст
        # Но для универсальности: если 1-12, считаем месяцем (01 = январь), если 13-53 - неделей
        match_numeric = re.match(r'(\d{4})[-_](\d{1,2})', period_str)
        if match_numeric:
            year = int(match_numeric.group(1))
            num = int(match_numeric.group(2))
            if 1 <= num <= 12:
                return (year, num, 0)  # 0 = месяц (01-12 это месяцы)
            elif 13 <= num <= 53:
                return (year, num, 1)  # 1 = неделя
        
        # Если ничего не подошло, возвращаем (0, 0, 0)
        return (0, 0, 0)
    except:
        return (0, 0, 0)

# Обратная совместимость
def parse_year_month(year_month_str):
    """Устаревшая функция, использует parse_period"""
    result = parse_period(year_month_str)
    return (result[0], result[1])

# Функция для цветового форматирования матрицы (градиент красный-желтый-зеленый)
def color_gradient(val, min_val, max_val, mean_val, is_diagonal=False):
    """Применяет четкий градиент от красного (минимум) через желтый (среднее) к зеленому (максимум)
    Если is_diagonal=True, возвращает белый фон без цвета"""
    # Диагональные значения (сама когорта) - без цвета, жирный шрифт, по центру
    if is_diagonal:
        return 'background-color: white; color: black; font-weight: bold; text-align: center'
    
    if pd.isna(val) or val == 0:
        return 'background-color: white; color: black; text-align: center'
    
    # Если значение меньше или равно среднему - градиент от красного к желтому
    if val <= mean_val:
        # Градиент от красного (255,0,0) к желтому (255,255,0)
        if mean_val == min_val:
            ratio = 1.0  # Все значения равны минимуму, делаем желтым
        else:
            ratio = (val - min_val) / (mean_val - min_val)
            ratio = max(0, min(1, ratio))  # Ограничиваем от 0 до 1
        
        # Красный -> Желтый: R=255 постоянный, G растет от 0 до 255, B=0 постоянный
        r = 255
        g = int(255 * ratio)  # от 0 до 255
        b = 0
    else:
        # Градиент от желтого (255,255,0) к зеленому (0,255,0)
        if max_val == mean_val:
            ratio = 1.0  # Все значения равны среднему, делаем зеленым
        else:
            ratio = (val - mean_val) / (max_val - mean_val)
            ratio = max(0, min(1, ratio))  # Ограничиваем от 0 до 1
        
        # Желтый -> Зеленый: R убывает от 255 до 0, G=255 постоянный, B=0 постоянный
        r = int(255 * (1 - ratio))  # от 255 до 0
        g = 255
        b = 0
    
    # Всегда используем чёрный цвет текста и выравнивание по центру
    return f'background-color: rgb({r},{g},{b}); color: black; text-align: center'

def apply_matrix_color_gradient(df, hide_zeros=False, horizontal_dynamics=False, hide_before_diagonal=False):
    """Применяет цветовое форматирование к матрице
    Диагональные значения (сама когорта) отображаются без цвета, жирным шрифтом
    
    Parameters:
    - df: DataFrame для форматирования
    - hide_zeros: если True, нулевые значения скрываются (пустая строка)
    - horizontal_dynamics: если True, градиент рассчитывается по каждой строке отдельно
    - hide_before_diagonal: если True, скрываются все значения до диагонали (для горизонтальной динамики)
    """
    # Получаем индексы периодов для определения порядка
    period_indices = {period: idx for idx, period in enumerate(df.index)}
    
    # Если нужно скрывать нули или значения до диагонали, заменяем значения на пустую строку перед форматированием
    df_display = df.copy()
    if hide_zeros or hide_before_diagonal:
        for row_name in df_display.index:
            row_idx = period_indices.get(row_name, 0)
            for col_name in df_display.columns:
                col_idx = period_indices.get(col_name, 0)
                is_diagonal = (row_name == col_name)
                
                # Скрываем значения до диагонали (если период меньше когорты)
                if hide_before_diagonal and not is_diagonal and col_idx < row_idx:
                    df_display.loc[row_name, col_name] = ''
                # Скрываем нулевые значения
                elif hide_zeros and not is_diagonal and (pd.isna(df_display.loc[row_name, col_name]) or df_display.loc[row_name, col_name] == 0):
                    df_display.loc[row_name, col_name] = ''
    
    # Применяем форматирование с учетом позиции (диагональные значения без цвета)
    def format_with_diagonal(x):
        """Применяет форматирование с учетом диагонали"""
        result = pd.DataFrame(index=x.index, columns=x.columns, dtype=object)
        
        # Получаем индексы для определения порядка в функции форматирования
        period_indices_format = {period: idx for idx, period in enumerate(x.index)}
        
        for row_name in x.index:
            row_idx_format = period_indices_format.get(row_name, 0)
            
            # Для горизонтальной динамики рассчитываем min/max/mean для каждой строки отдельно
            if horizontal_dynamics:
                row_values = []
                for col_name in x.columns:
                    col_idx_format = period_indices_format.get(col_name, 0)
                    # Учитываем только значения после диагонали (если hide_before_diagonal включен) или все недиагональные
                    if row_name != col_name and (not hide_before_diagonal or col_idx_format >= row_idx_format):
                        val = x.loc[row_name, col_name]
                        val_for_calc = 0 if (val == '' or pd.isna(val)) else val
                        if val_for_calc != 0:
                            row_values.append(val_for_calc)
                
                if row_values:
                    row_min = min(row_values)
                    row_max = max(row_values)
                    row_mean = sum(row_values) / len(row_values)
                else:
                    row_min = 0
                    row_max = 0
                    row_mean = 0
            else:
                # Глобальный расчет для всей таблицы (исключая диагональ)
                non_diagonal_values = []
                for r_name in x.index:
                    for c_name in x.columns:
                        if r_name != c_name:
                            val = x.loc[r_name, c_name]
                            # Преобразуем значение в число, если это строка с процентом
                            if isinstance(val, str):
                                # Пытаемся извлечь число из строки типа "45.7%"
                                try:
                                    val_for_calc = float(val.replace('%', '').strip())
                                except (ValueError, AttributeError):
                                    val_for_calc = 0
                            else:
                                val_for_calc = 0 if (val == '' or pd.isna(val)) else float(val)
                            
                            if val_for_calc != 0:
                                non_diagonal_values.append(val_for_calc)
                
                if non_diagonal_values:
                    row_min = min(non_diagonal_values)
                    row_max = max(non_diagonal_values)
                    row_mean = sum(non_diagonal_values) / len(non_diagonal_values)
                else:
                    row_min = 0
                    row_max = 0
                    row_mean = 0
            
            for col_name in x.columns:
                val = x.loc[row_name, col_name]
                is_diagonal = (row_name == col_name)
                
                # Если значение пустое (скрытое), применяем прозрачный стиль
                col_idx_display = period_indices.get(col_name, 0)
                row_idx_display = period_indices.get(row_name, 0)
                
                is_hidden = (
                    (hide_zeros and not is_diagonal and (val == '' or pd.isna(val) or val == 0)) or
                    (hide_before_diagonal and not is_diagonal and col_idx_display < row_idx_display)
                )
                
                if is_hidden:
                    result.loc[row_name, col_name] = 'background-color: white; color: white; text-align: center'
                else:
                    # Преобразуем значение для расчета цвета
                    # Если значение - строка с процентом, извлекаем число
                    if isinstance(val, str) and '%' in val:
                        try:
                            val_for_color = float(val.replace('%', '').strip())
                        except (ValueError, AttributeError):
                            val_for_color = 0
                    else:
                        val_for_color = 0 if (val == '' or pd.isna(val)) else float(val) if not isinstance(val, str) else 0
                    
                    gradient_style = color_gradient(val_for_color, row_min, row_max, row_mean, is_diagonal)
                    # Добавляем выравнивание по центру (если еще не добавлено)
                    if 'text-align' not in gradient_style:
                        gradient_style += '; text-align: center'
                    result.loc[row_name, col_name] = gradient_style
        return result
    
    styled_df = df_display.style.apply(format_with_diagonal, axis=None)
    
    return styled_df

def apply_excel_color_formatting(worksheet, df, hide_zeros=False):
    """Применяет цветовое форматирование к Excel файлу
    Parameters:
    - worksheet: лист Excel
    - df: DataFrame для форматирования
    - hide_zeros: если True, нулевые значения скрываются (пустая ячейка)
    """
    min_val = df.min().min()
    max_val = df.max().max()
    mean_val = df.mean().mean()
    
    def get_rgb_color(val, min_val, max_val, mean_val, is_diagonal=False):
        """Возвращает RGB цвет для значения - четкий градиент от красного к желтому, от желтого к зеленому"""
        # Диагональные значения - белый фон
        if is_diagonal:
            return (255, 255, 255)  # белый
        
        if pd.isna(val) or val == 0:
            return (255, 255, 255)  # белый
        
        # Если значение меньше или равно среднему - градиент от красного к желтому
        if val <= mean_val:
            # Градиент от красного (255,0,0) к желтому (255,255,0)
            if mean_val == min_val:
                ratio = 1.0  # Все значения равны минимуму, делаем желтым
            else:
                ratio = (val - min_val) / (mean_val - min_val)
                ratio = max(0, min(1, ratio))  # Ограничиваем от 0 до 1
            
            # Красный -> Желтый: R=255 постоянный, G растет от 0 до 255, B=0 постоянный
            r = 255
            g = int(255 * ratio)  # от 0 до 255
            b = 0
        else:
            # Градиент от желтого (255,255,0) к зеленому (0,255,0)
            if max_val == mean_val:
                ratio = 1.0  # Все значения равны среднему, делаем зеленым
            else:
                ratio = (val - mean_val) / (max_val - mean_val)
                ratio = max(0, min(1, ratio))  # Ограничиваем от 0 до 1
            
            # Желтый -> Зеленый: R убывает от 255 до 0, G=255 постоянный, B=0 постоянный
            r = int(255 * (1 - ratio))  # от 255 до 0
            g = 255
            b = 0
        
        return (r, g, b)
    
    # Применяем форматирование к данным (начиная со строки 2, т.к. строка 1 - заголовки)
    period_indices_excel = {period: idx for idx, period in enumerate(df.index)}
    
    # Определяем, на какой строке начинаются данные (обычно строка 2, если есть заголовок индекса)
    # Если индекс имеет имя, то заголовок в строке 1, данные начинаются со строки 2
    start_row = 2  # Начальная строка с данными (строка 1 - заголовки столбцов и индекса)
    
    for row_idx, period in enumerate(df.index, start=start_row):
        for col_idx, col_period in enumerate(df.columns, start=2):  # Столбец 1 - индекс, данные с столбца 2
            cell = worksheet.cell(row=row_idx, column=col_idx)
            value = df.loc[period, col_period]
            
            # Проверяем, является ли это диагональю
            is_diagonal = (period == col_period)
            
            if is_diagonal:
                # Диагональ - белый фон, жирный шрифт
                r, g, b = get_rgb_color(value, min_val, max_val, mean_val, is_diagonal=True)
                hex_color = f"{r:02X}{g:02X}{b:02X}"
                cell.fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
                cell.font = Font(color="000000", bold=True)  # чёрный текст, жирный
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif not pd.isna(value) and value != 0:
                r, g, b = get_rgb_color(value, min_val, max_val, mean_val, is_diagonal=False)
                # Формат RGB для openpyxl: RRGGBB
                hex_color = f"{r:02X}{g:02X}{b:02X}"
                cell.fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
                cell.font = Font(color="000000")  # чёрный текст
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                # Нулевые значения или пустые
                if hide_zeros and not is_diagonal:
                    # Скрываем нули (пустая ячейка)
                    cell.value = ""
                    cell.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                    cell.font = Font(color="FFFFFF")  # белый текст на белом фоне
                else:
                    # Показываем нули
                    cell.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                    cell.font = Font(color="000000")
                cell.alignment = Alignment(horizontal="center", vertical="center")

def apply_excel_cohort_formatting(worksheet, df, sorted_periods):
    """Применяет цветовое форматирование с горизонтальной динамикой к Excel файлу для таблицы когорт"""
    period_indices = {period: idx for idx, period in enumerate(sorted_periods)}
    
    # Для горизонтальной динамики рассчитываем min/max/mean для каждой строки отдельно
    def get_row_stats(row_period):
        row_idx = period_indices.get(row_period, 0)
        row_values = []
        for col_period in df.columns:
            col_idx = period_indices.get(col_period, 0)
            # Учитываем только значения после диагонали
            if row_period != col_period and col_idx >= row_idx:
                val = df.loc[row_period, col_period]
                if not pd.isna(val) and val > 0:
                    row_values.append(val)
        if row_values:
            return min(row_values), max(row_values), sum(row_values) / len(row_values)
        return 0, 0, 0
    
    def get_rgb_color_cohort(val, min_val, max_val, mean_val, is_diagonal=False):
        """Возвращает RGB цвет для значения"""
        if is_diagonal:
            return (255, 255, 255)  # белый для диагонали
        
        if pd.isna(val) or val == 0:
            return (255, 255, 255)  # белый
        
        if val <= mean_val:
            if mean_val == min_val:
                ratio = 1.0
            else:
                ratio = (val - min_val) / (mean_val - min_val)
                ratio = max(0, min(1, ratio))
            r = 255
            g = int(255 * ratio)
            b = 0
        else:
            if max_val == mean_val:
                ratio = 1.0
            else:
                ratio = (val - mean_val) / (max_val - mean_val)
                ratio = max(0, min(1, ratio))
            r = int(255 * (1 - ratio))
            g = 255
            b = 0
        return (r, g, b)
    
    start_row = 2
    for row_idx, period in enumerate(df.index, start=start_row):
        row_period_idx = period_indices.get(period, 0)
        row_min, row_max, row_mean = get_row_stats(period)
        
        for col_idx, col_period in enumerate(df.columns, start=2):
            col_period_idx = period_indices.get(col_period, 0)
            cell = worksheet.cell(row=row_idx, column=col_idx)
            value = df.loc[period, col_period]
            is_diagonal = (period == col_period)
            
            # Скрываем значения до диагонали
            if not is_diagonal and col_period_idx < row_period_idx:
                cell.value = ""
                cell.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                cell.font = Font(color="FFFFFF")  # белый текст на белом фоне
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif is_diagonal:
                # Диагональ - белый фон, жирный шрифт
                r, g, b = get_rgb_color_cohort(value, row_min, row_max, row_mean, is_diagonal=True)
                hex_color = f"{r:02X}{g:02X}{b:02X}"
                cell.fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
                cell.font = Font(color="000000", bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                # Форматируем как целое число
                if cell.value is not None and not isinstance(cell.value, str):
                    cell.number_format = '0'
            elif not pd.isna(value) and value > 0:
                r, g, b = get_rgb_color_cohort(value, row_min, row_max, row_mean, is_diagonal=False)
                hex_color = f"{r:02X}{g:02X}{b:02X}"
                cell.fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
                cell.font = Font(color="000000")
                cell.alignment = Alignment(horizontal="center", vertical="center")
                # Форматируем как целое число
                if cell.value is not None and not isinstance(cell.value, str):
                    cell.number_format = '0'
            else:
                cell.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                cell.font = Font(color="000000")
                cell.alignment = Alignment(horizontal="center", vertical="center")

def apply_excel_percent_formatting(worksheet, df, sorted_periods):
    """Применяет цветовое форматирование и форматирование процентов к Excel файлу для таблицы накопления в %"""
    period_indices = {period: idx for idx, period in enumerate(sorted_periods)}
    
    # Для горизонтальной динамики рассчитываем min/max/mean для каждой строки отдельно
    def get_row_stats(row_period):
        row_idx = period_indices.get(row_period, 0)
        row_values = []
        for col_period in df.columns:
            col_idx = period_indices.get(col_period, 0)
            if row_period != col_period and col_idx >= row_idx:
                val = df.loc[row_period, col_period]
                if not pd.isna(val) and val > 0:
                    row_values.append(val)
        if row_values:
            return min(row_values), max(row_values), sum(row_values) / len(row_values)
        return 0, 0, 0
    
    def get_rgb_color_percent(val, min_val, max_val, mean_val, is_diagonal=False):
        """Возвращает RGB цвет для значения"""
        if is_diagonal:
            return (255, 255, 255)  # белый для диагонали
        
        if pd.isna(val) or val == 0:
            return (255, 255, 255)  # белый
        
        if val <= mean_val:
            if mean_val == min_val:
                ratio = 1.0
            else:
                ratio = (val - min_val) / (mean_val - min_val)
                ratio = max(0, min(1, ratio))
            r = 255
            g = int(255 * ratio)
            b = 0
        else:
            if max_val == mean_val:
                ratio = 1.0
            else:
                ratio = (val - mean_val) / (max_val - mean_val)
                ratio = max(0, min(1, ratio))
            r = int(255 * (1 - ratio))
            g = 255
            b = 0
        return (r, g, b)
    
    start_row = 2
    for row_idx, period in enumerate(df.index, start=start_row):
        row_period_idx = period_indices.get(period, 0)
        row_min, row_max, row_mean = get_row_stats(period)
        
        for col_idx, col_period in enumerate(df.columns, start=2):
            col_period_idx = period_indices.get(col_period, 0)
            cell = worksheet.cell(row=row_idx, column=col_idx)
            value = df.loc[period, col_period]
            is_diagonal = (period == col_period)
            
            # Скрываем значения до диагонали
            if not is_diagonal and col_period_idx < row_period_idx:
                cell.value = ""
                cell.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                cell.font = Font(color="FFFFFF")  # белый текст на белом фоне
            elif is_diagonal:
                # Диагональ - 100.0% (сохраняем как число 1.0, Excel покажет как 100%)
                cell.value = 1.0
                cell.number_format = '0.0%'  # Процентный формат Excel
                r, g, b = get_rgb_color_percent(100.0, row_min, row_max, row_mean, is_diagonal=True)
                hex_color = f"{r:02X}{g:02X}{b:02X}"
                cell.fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
                cell.font = Font(color="000000", bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif not pd.isna(value) and value > 0:
                # Сохраняем как число (value уже в процентах, конвертируем в долю для Excel)
                cell.value = value / 100.0  # Конвертируем проценты в долю (45.7 -> 0.457)
                cell.number_format = '0.0%'  # Процентный формат Excel
                r, g, b = get_rgb_color_percent(value, row_min, row_max, row_mean, is_diagonal=False)
                hex_color = f"{r:02X}{g:02X}{b:02X}"
                cell.fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
                cell.font = Font(color="000000")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.value = ""
                cell.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                cell.font = Font(color="FFFFFF")
                cell.alignment = Alignment(horizontal="center", vertical="center")

def apply_excel_inflow_formatting(worksheet, df, sorted_periods):
    """Применяет цветовое форматирование и форматирование процентов к Excel файлу для таблицы притока в %"""
    period_indices = {period: idx for idx, period in enumerate(sorted_periods)}
    
    # Для горизонтальной динамики рассчитываем min/max/mean для каждой строки отдельно
    def get_row_stats(row_period):
        row_idx = period_indices.get(row_period, 0)
        row_values = []
        for col_period in df.columns:
            col_idx = period_indices.get(col_period, 0)
            if row_period != col_period and col_idx >= row_idx:
                val = df.loc[row_period, col_period]
                if not pd.isna(val) and val > 0:
                    row_values.append(val)
        if row_values:
            return min(row_values), max(row_values), sum(row_values) / len(row_values)
        return 0, 0, 0
    
    def get_rgb_color_inflow(val, min_val, max_val, mean_val, is_diagonal=False):
        """Возвращает RGB цвет для значения"""
        if is_diagonal:
            return (255, 255, 255)  # белый для диагонали
        
        if pd.isna(val) or val == 0:
            return (255, 255, 255)  # белый
        
        if val <= mean_val:
            if mean_val == min_val:
                ratio = 1.0
            else:
                ratio = (val - min_val) / (mean_val - min_val)
                ratio = max(0, min(1, ratio))
            r = 255
            g = int(255 * ratio)
            b = 0
        else:
            if max_val == mean_val:
                ratio = 1.0
            else:
                ratio = (val - mean_val) / (max_val - mean_val)
                ratio = max(0, min(1, ratio))
            r = int(255 * (1 - ratio))
            g = 255
            b = 0
        return (r, g, b)
    
    start_row = 2
    for row_idx, period in enumerate(df.index, start=start_row):
        row_period_idx = period_indices.get(period, 0)
        row_min, row_max, row_mean = get_row_stats(period)
        
        for col_idx, col_period in enumerate(df.columns, start=2):
            col_period_idx = period_indices.get(col_period, 0)
            cell = worksheet.cell(row=row_idx, column=col_idx)
            value = df.loc[period, col_period]
            is_diagonal = (period == col_period)
            
            # Скрываем значения до диагонали
            if not is_diagonal and col_period_idx < row_period_idx:
                cell.value = ""
                cell.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                cell.font = Font(color="FFFFFF")  # белый текст на белом фоне
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif is_diagonal:
                # Диагональ - 0.0% (сохраняем как число 0.0, Excel покажет как 0.0%)
                cell.value = 0.0
                cell.number_format = '0.0%'  # Процентный формат Excel
                r, g, b = get_rgb_color_inflow(0.0, row_min, row_max, row_mean, is_diagonal=True)
                hex_color = f"{r:02X}{g:02X}{b:02X}"
                cell.fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
                cell.font = Font(color="000000", bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif not pd.isna(value) and value > 0:
                # Сохраняем как число (value уже в процентах, конвертируем в долю для Excel)
                cell.value = value / 100.0  # Конвертируем проценты в долю (45.7 -> 0.457)
                cell.number_format = '0.0%'  # Процентный формат Excel
                r, g, b = get_rgb_color_inflow(value, row_min, row_max, row_mean, is_diagonal=False)
                hex_color = f"{r:02X}{g:02X}{b:02X}"
                cell.fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
                cell.font = Font(color="000000")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.value = ""
                cell.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                cell.font = Font(color="FFFFFF")
                cell.alignment = Alignment(horizontal="center", vertical="center")

# Функция построения когортной матрицы
def build_cohort_matrix(df, year_month_col, client_col, value_type='clients'):
    """
    Строит когортную матрицу по периоду "Год-месяц"
    
    Parameters:
    - df: DataFrame с данными
    - year_month_col: название столбца с годом-месяцем
    - client_col: название столбца с кодом клиента
    - value_type: тип значений в матрице ('clients' - уникальные клиенты, 'count' - количество записей)
    """
    # Получаем уникальные периоды и сортируем их
    unique_periods = df[year_month_col].dropna().unique()
    
    # Сортируем периоды по году и номеру периода (месяц или неделя)
    periods_with_sort = [(period, parse_period(str(period).strip())) for period in unique_periods]
    
    # Сортируем: сначала по году, потом по типу (месяцы сначала), потом по номеру
    # Периоды с (0, 0, 0) будут в начале, поэтому фильтруем их
    valid_periods = [(p, parsed) for p, parsed in periods_with_sort if parsed != (0, 0, 0)]
    invalid_periods = [p for p, parsed in periods_with_sort if parsed == (0, 0, 0)]
    
    if valid_periods:
        valid_periods.sort(key=lambda x: (x[1][0], x[1][2], x[1][1]))  # (year, type, number)
        sorted_periods = [period[0] for period in valid_periods]
        
        # Добавляем нераспознанные периоды в конец (если есть)
        if invalid_periods:
            sorted_periods.extend(sorted(invalid_periods))
    else:
        # Если все периоды не распознаны, используем просто сортировку по строке
        sorted_periods = sorted([str(p) for p in unique_periods])
    
    # Оптимизация: предварительно группируем данные по периодам
    # Создаем словарь: период -> множество клиентов
    period_clients = {}
    for period in sorted_periods:
        period_data = df[df[year_month_col] == period]
        if value_type == 'clients':
            period_clients[period] = set(period_data[client_col].dropna().unique())
        else:
            # Для count просто сохраняем количество
            period_clients[period] = len(period_data)
    
    # Создаем матрицу пересечений клиентов
    matrix_intersection = pd.DataFrame(
        index=sorted_periods,
        columns=sorted_periods,
        dtype=int
    )
    
    # Заполняем матрицу используя предварительно вычисленные множества
    for row_period in sorted_periods:
        for col_period in sorted_periods:
            if row_period == col_period:
                # Диагональ - клиенты в этом периоде
                if value_type == 'clients':
                    matrix_intersection.loc[row_period, col_period] = len(period_clients[row_period])
                else:
                    matrix_intersection.loc[row_period, col_period] = period_clients[row_period]
            else:
                # Пересечение клиентов между двумя периодами
                if value_type == 'clients':
                    clients_row = period_clients[row_period]
                    clients_col = period_clients[col_period]
                    intersection = len(clients_row & clients_col)
                    matrix_intersection.loc[row_period, col_period] = intersection
                else:
                    # Для count это не имеет смысла, но оставляем для совместимости
                    matrix_intersection.loc[row_period, col_period] = 0
    
    return matrix_intersection, sorted_periods

# Функция построения матрицы накопления возврата
def build_accumulation_matrix(df, year_month_col, client_col, sorted_periods):
    """
    Строит матрицу накопления возврата клиентов
    Накопление идет только с периода СЛЕДУЮЩЕГО за периодом когорты (без самого периода когорты)
    
    Parameters:
    - df: DataFrame с данными
    - year_month_col: название столбца с годом-месяцем
    - client_col: название столбца с кодом клиента
    - sorted_periods: отсортированный список периодов
    
    Returns:
    - matrix_accumulation: матрица накопления уникальных клиентов
    """
    matrix_accumulation = pd.DataFrame(
        index=sorted_periods,
        columns=sorted_periods,
        dtype=int
    )
    
    # Оптимизация: предварительно создаем словарь период -> множество клиентов
    period_clients_dict = {}
    for period in sorted_periods:
        period_data = df[df[year_month_col] == period]
        period_clients_dict[period] = set(period_data[client_col].dropna().unique())
    
    # Получаем индекс каждого периода для определения порядка
    period_indices = {period: idx for idx, period in enumerate(sorted_periods)}
    
    for row_period in sorted_periods:
        row_idx = period_indices[row_period]
        
        # Получаем множество клиентов этой когорты (в первом периоде когорты)
        cohort_clients = period_clients_dict[row_period]
        
        # Предварительно вычисляем накопление для всех последующих периодов
        accumulated_clients_by_period = {}
        current_accumulated = set()
        
        for col_idx in range(row_idx, len(sorted_periods)):
            col_period = sorted_periods[col_idx]
            
            if col_idx == row_idx:
                # Диагональ - клиенты в первом периоде когорты
                matrix_accumulation.loc[row_period, col_period] = len(cohort_clients)
                accumulated_clients_by_period[col_period] = set(cohort_clients)
            elif col_idx > row_idx:
                # Добавляем клиентов из текущего периода к накопленным
                period_clients = period_clients_dict[col_period]
                cohort_period_clients = period_clients & cohort_clients
                current_accumulated.update(cohort_period_clients)
                accumulated_clients_by_period[col_period] = set(current_accumulated)
                matrix_accumulation.loc[row_period, col_period] = len(current_accumulated)
        
        # Заполняем нулями периоды до начала когорты
        for col_idx in range(row_idx):
            col_period = sorted_periods[col_idx]
            matrix_accumulation.loc[row_period, col_period] = 0
    
    return matrix_accumulation

# Функция построения матрицы накопления возврата в процентах
# Функции для получения кодов клиентов из матриц
def get_cohort_clients(df, year_month_col, client_col, cohort_period, target_period, period_clients_cache=None):
    """Получает коды клиентов из когорты, которые были в целевом периоде"""
    if period_clients_cache:
        clients_in_cohort = period_clients_cache.get(cohort_period, set())
        clients_in_period = period_clients_cache.get(target_period, set())
    else:
        clients_in_cohort = set(df[df[year_month_col] == cohort_period][client_col].dropna().unique())
        clients_in_period = set(df[df[year_month_col] == target_period][client_col].dropna().unique())
    return sorted(list(clients_in_cohort & clients_in_period))

def get_accumulation_clients(df, year_month_col, client_col, sorted_periods, cohort_period, target_period, period_clients_cache=None):
    """Получает накопленные коды клиентов из когорты до целевого периода включительно (без самого периода когорты)"""
    period_indices = {period: idx for idx, period in enumerate(sorted_periods)}
    cohort_idx = period_indices.get(cohort_period, -1)
    target_idx = period_indices.get(target_period, -1)
    
    if cohort_idx < 0 or target_idx < 0 or target_idx <= cohort_idx:
        return []
    
    # Получаем множество клиентов этой когорты
    if period_clients_cache:
        cohort_clients = period_clients_cache.get(cohort_period, set())
    else:
        cohort_clients = set(df[df[year_month_col] == cohort_period][client_col].dropna().unique())
    
    # Находим всех клиентов когорты, которые вернулись в любом периоде от следующего после когорты до целевого включительно
    returned_clients = set()
    for period in sorted_periods[cohort_idx + 1:target_idx + 1]:
        if period_clients_cache:
            period_clients = period_clients_cache.get(period, set())
        else:
            period_clients = set(df[df[year_month_col] == period][client_col].dropna().unique())
        returned_clients.update(cohort_clients & period_clients)
    
    return sorted(list(returned_clients))

def get_client_cohorts(df, year_month_col, client_col, sorted_periods):
    """
    Определяет когорту для каждого клиента (первый период появления)
    Returns: словарь {client: cohort_period}
    Оптимизированная версия с использованием groupby
    """
    # Создаем словарь индексов периодов для быстрого поиска
    period_indices = {period: idx for idx, period in enumerate(sorted_periods)}
    
    # Фильтруем только валидные периоды и клиентов
    df_filtered = df[[year_month_col, client_col]].dropna()
    
    # Группируем по клиентам и находим минимальный период (первый по индексу)
    # Используем groupby для эффективности
    client_first_periods = df_filtered.groupby(client_col)[year_month_col].min()
    
    # Преобразуем в словарь, фильтруя только валидные периоды
    client_cohorts = {}
    for client, first_period in client_first_periods.items():
        if first_period in period_indices:
            client_cohorts[client] = first_period
    
    return client_cohorts

def get_churn_clients(df, year_month_col, client_col, sorted_periods, cohort_period, period_clients_cache=None, client_cohorts_cache=None):
    """
    Получает коды клиентов оттока из когорты (те, кто не вернулся ни разу после периода когорты)
    Важно: когорта определяется как первый период появления клиента, а не просто все клиенты из периода
    Это гарантирует, что каждый клиент попадает только в одну когорту, устраняя дубли
    """
    period_indices = {period: idx for idx, period in enumerate(sorted_periods)}
    cohort_idx = period_indices.get(cohort_period, -1)
    
    if cohort_idx < 0:
        return []
    
    # Получаем когорты клиентов (если кэш не передан, вычисляем)
    if client_cohorts_cache is None:
        client_cohorts_cache = get_client_cohorts(df, year_month_col, client_col, sorted_periods)
    
    # Получаем множество клиентов, для которых указанный период является их когортой (первым появлением)
    cohort_clients = set()
    for client, client_cohort in client_cohorts_cache.items():
        if client_cohort == cohort_period:
            cohort_clients.add(client)
    
    # Если когорта пустая, возвращаем пустой список
    if not cohort_clients:
        return []
    
    # Находим всех клиентов когорты, которые вернулись хотя бы раз в любом периоде после когорты
    returned_clients = set()
    for period in sorted_periods[cohort_idx + 1:]:
        if period_clients_cache:
            period_clients = period_clients_cache.get(period, set())
        else:
            period_clients = set(df[df[year_month_col] == period][client_col].dropna().unique())
        returned_clients.update(cohort_clients & period_clients)
    
    # Отток = клиенты когорты - вернувшиеся клиенты
    churn_clients = cohort_clients - returned_clients
    return sorted(list(churn_clients))

def build_churn_table(df, year_month_col, client_col, sorted_periods, cohort_matrix, accumulation_matrix, accumulation_percent_matrix, client_cohorts_cache=None, period_clients_cache=None):
    """
    Строит таблицу оттока клиентов для всех когорт
    Использует исходную логику: когорта = все клиенты из периода (не только первый период появления)
    """
    churn_data = []
    
    # Оптимизация: создаём period_indices один раз вне цикла
    period_indices = {period: idx for idx, period in enumerate(sorted_periods)}
    last_period = sorted_periods[-1]
    last_period_idx = period_indices[last_period]
    
    for cohort_period in sorted_periods:
        # 1. Когорта
        cohort = cohort_period
        
        # 2. Кол-во клиентов когорты (из матрицы когорт - все клиенты из этого периода)
        cohort_size = cohort_matrix.loc[cohort_period, cohort_period]
        
        # 3. Накопительное кол-во возврата за весь период
        # Берем последний столбец (последний период) для этой когорты
        cohort_idx = period_indices[cohort_period]
        
        if last_period_idx > cohort_idx:
            # Если есть периоды после когорты, берем значение из матрицы накопления
            total_returned = accumulation_matrix.loc[cohort_period, last_period]
        else:
            # Если это последняя когорта, возврат = 0
            total_returned = 0
        
        # 4. Накопительный % возврата за весь период
        if cohort_size > 0:
            total_returned_percent = (total_returned / cohort_size) * 100
        else:
            total_returned_percent = 0
        
        # 5. Отток кол-во = клиенты когорты - вернувшиеся
        churn_count = int(cohort_size - total_returned)
        
        # 6. Отток % = (отток / размер когорты) * 100
        if cohort_size > 0:
            churn_percent = (churn_count / cohort_size) * 100
        else:
            churn_percent = 0
        
        churn_data.append({
            'Когорта': cohort,
            'Кол-во клиентов когорты': int(cohort_size),
            'Накопительное кол-во возврата': int(total_returned),
            'Накопительный % возврата': total_returned_percent,
            'Отток кол-во': churn_count,
            'Отток %': churn_percent
        })
    
    churn_df = pd.DataFrame(churn_data)
    return churn_df

def get_inflow_clients(df, year_month_col, client_col, sorted_periods, cohort_period, target_period, period_clients_cache=None):
    """Получает коды клиентов из когорты, которые вернулись именно в целевом периоде (новый приток)"""
    period_indices = {period: idx for idx, period in enumerate(sorted_periods)}
    cohort_idx = period_indices.get(cohort_period, -1)
    target_idx = period_indices.get(target_period, -1)
    
    if cohort_idx < 0 or target_idx < 0 or target_idx <= cohort_idx:
        return []
    
    # Получаем множество клиентов этой когорты
    if period_clients_cache:
        cohort_clients = period_clients_cache.get(cohort_period, set())
    else:
        cohort_clients = set(df[df[year_month_col] == cohort_period][client_col].dropna().unique())
    
    # Клиенты, которые вернулись в целевом периоде
    if period_clients_cache:
        target_period_clients = period_clients_cache.get(target_period, set())
    else:
        target_period_clients = set(df[df[year_month_col] == target_period][client_col].dropna().unique())
    returned_in_target = cohort_clients & target_period_clients
    
    # Если это первый период после когорты, возвращаем всех вернувшихся
    if target_idx == cohort_idx + 1:
        return sorted(list(returned_in_target))
    
    # Иначе исключаем тех, кто уже вернулся ранее
    prev_periods_clients = set()
    for period in sorted_periods[cohort_idx + 1:target_idx]:
        if period_clients_cache:
            period_clients = period_clients_cache.get(period, set())
        else:
            period_clients = set(df[df[year_month_col] == period][client_col].dropna().unique())
        prev_periods_clients.update(cohort_clients & period_clients)
    
    # Новые возвраты в целевом периоде (не возвращались ранее)
    new_returns = returned_in_target - prev_periods_clients
    return sorted(list(new_returns))

def build_inflow_matrix(accumulation_percent_matrix):
    """
    Строит матрицу притока возврата в процентах
    Показывает прирост уникальных клиентов когорты между периодами
    
    Parameters:
    - accumulation_percent_matrix: матрица накопления в процентах
    
    Returns:
    - inflow_matrix: матрица притока в процентах
    """
    inflow_matrix = pd.DataFrame(
        index=accumulation_percent_matrix.index,
        columns=accumulation_percent_matrix.columns,
        dtype=float
    )
    
    # Получаем индексы периодов для определения порядка
    period_indices = {period: idx for idx, period in enumerate(accumulation_percent_matrix.index)}
    
    for row_period in accumulation_percent_matrix.index:
        row_idx = period_indices.get(row_period, 0)
        
        for col_period in accumulation_percent_matrix.columns:
            col_idx = period_indices.get(col_period, 0)
            
            # Диагональ = 0%
            if row_idx == col_idx:
                inflow_matrix.loc[row_period, col_period] = 0.0
            elif col_idx < row_idx:
                # До диагонали = 0
                inflow_matrix.loc[row_period, col_period] = 0.0
            else:
                # Первый столбец после диагонали = значение из матрицы накопления
                if col_idx == row_idx + 1:
                    inflow_matrix.loc[row_period, col_period] = accumulation_percent_matrix.loc[row_period, col_period]
                else:
                    # Остальные столбцы = разница между текущим и предыдущим значением
                    current_val = accumulation_percent_matrix.loc[row_period, col_period]
                    # Находим предыдущий период
                    prev_period = accumulation_percent_matrix.columns[col_idx - 1]
                    prev_val = accumulation_percent_matrix.loc[row_period, prev_period]
                    inflow_matrix.loc[row_period, col_period] = current_val - prev_val
    
    return inflow_matrix

def build_accumulation_percent_matrix(accumulation_matrix, cohort_matrix):
    """
    Строит матрицу накопления возврата в процентах
    Доля накопления количества клиентов от количества клиентов в когорте
    
    Parameters:
    - accumulation_matrix: матрица накопления (абсолютные значения)
    - cohort_matrix: исходная когортная матрица (для получения количества клиентов в когорте)
    
    Returns:
    - matrix_percent: матрица в процентах
    """
    matrix_percent = pd.DataFrame(
        index=accumulation_matrix.index,
        columns=accumulation_matrix.columns,
        dtype=float
    )
    
    # Получаем индексы периодов для определения порядка
    period_indices = {period: idx for idx, period in enumerate(accumulation_matrix.index)}
    
    for row_period in accumulation_matrix.index:
        row_idx = period_indices.get(row_period, 0)
        
        # Количество клиентов в когорте (диагональ)
        cohort_size = cohort_matrix.loc[row_period, row_period]
        
        for col_period in accumulation_matrix.columns:
            col_idx = period_indices.get(col_period, 0)
            
            if col_idx < row_idx:
                # Период до начала когорты
                matrix_percent.loc[row_period, col_period] = 0
            elif col_idx == row_idx:
                # Диагональ - 100% (все клиенты когорты)
                matrix_percent.loc[row_period, col_period] = 100.0 if cohort_size > 0 else 0
            else:
                # Процент накопления: (накопление / размер когорты) * 100
                accumulation_value = accumulation_matrix.loc[row_period, col_period]
                if cohort_size > 0:
                    percent = (accumulation_value / cohort_size) * 100
                    matrix_percent.loc[row_period, col_period] = percent
                else:
                    matrix_percent.loc[row_period, col_period] = 0
    
    return matrix_percent

# Функция загрузки Excel файла
# Создаем колонки для выравнивания заголовков на одном уровне
col_header_left, col_header_right = st.columns([1, 1])

with col_header_left:
    st.header("📊 Когортный анализ")

with col_header_right:
    st.subheader("📋 Шаблон загрузки данных из Qlik")

# Блок шаблона Qlik - инструкции слева, изображение и загрузчик справа
col_template_instructions, col_template_image = st.columns([1, 1])

with col_template_image:
    # Пытаемся найти скриншот шаблона Qlik (изображение вверху)
    qlik_image_paths = [
        'Qlik.png',
        'Qlik.jpg',
        'Qlik.jpeg',
        'qlik_template.png',
        'qlik_template.jpg',
        'qlik_template.jpeg',
        'шаблон_qlik.png',
        'шаблон_qlik.jpg',
        'шаблон_qlik.jpeg',
        'qlik.png',
        'qlik.jpg',
        'qlik.jpeg'
    ]
    image_found = False
    for img_path in qlik_image_paths:
        if os.path.exists(img_path):
            st.image(img_path, use_container_width=True)
            image_found = True
            break
    if not image_found:
        st.info("📸 Поместите скриншот шаблона загрузки данных из Qlik в папку проекта с одним из имён: Qlik.png, qlik_template.png, шаблон_qlik.png или qlik.png")
    
    # Загрузчик Excel файла прямо под изображением
    uploaded_file = st.file_uploader(
        "Выберите Excel файл для загрузки",
        type=['xlsx', 'xls'],
        help="Поддерживаются файлы формата .xlsx и .xls"
    )

with col_template_instructions:
    # Текст инструкций
    st.markdown("""
    1. Зайдите в Qlik, анализ чеков.
    
    2. Отберите необходимую категорию и уровни товара.
    
    3. Отберите анализируемый период.
    
    4. Зайдите на лист "Конструктор" и выведите отчёт по шаблону справа.
    
    Настройте фильтрами построение динамики когорт: Год-Месяц или Год-Неделя.
    
    5. Скачайте документ в Qlik и загрузите в ячейку справа.
    """)

if uploaded_file is not None:
    try:
        # Загрузка Excel файла
        if uploaded_file.name.endswith('.xlsx'):
            df = pd.read_excel(uploaded_file, engine='openpyxl')
        else:
            df = pd.read_excel(uploaded_file, engine='xlrd')
        
        # Сохранение данных в session state
        # Проверяем, новый ли это файл
        is_new_file = (
            st.session_state.uploaded_data is None or 
            st.session_state.uploaded_data.name != uploaded_file.name
        )
        
        st.session_state.uploaded_data = uploaded_file
        st.session_state.df = df
        
        # Очищаем старую информацию только при загрузке нового файла
        if is_new_file:
            st.session_state.cohort_info = None
            st.session_state.cohort_matrix = None
            st.session_state.sorted_periods = None
            st.session_state.year_month_col = None
            st.session_state.client_col = None
        
        # Построение когортной матрицы
        # Уменьшаем отступ перед блоком матриц
        st.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True)
        
        # Определяем столбцы автоматически
        expected_columns = {
            'Год-месяц': 'Год-месяц',
            'Год-Неделя': 'Год-Неделя',
            'Год-неделя': 'Год-неделя',
            'Год-Месяц': 'Год-Месяц',
            'Код клиента': 'Код клиента'
        }
        
        # Проверяем наличие ожидаемых столбцов
        year_month_col = None
        client_col = None
        
        # Ищем столбец с периодом (год-месяц или год-неделя)
        for col in df.columns:
            col_lower = str(col).lower()
            if 'год' in col_lower and ('месяц' in col_lower or 'неделя' in col_lower or 'неделя' in col_lower):
                year_month_col = col
                break
        
        # Ищем столбец с кодом клиента
        for col in df.columns:
            col_lower = str(col).lower()
            if 'код' in col_lower and 'клиент' in col_lower:
                client_col = col
                break
        
        # Если столбцы не найдены, показываем ошибку
        if year_month_col is None:
            st.error("❌ Не найден столбец с периодом (Год-месяц или Год-Неделя). Убедитесь, что в файле есть столбец с названием, содержащим 'Год' и 'месяц' или 'неделя'.")
            st.stop()
        
        if client_col is None:
            st.error("❌ Не найден столбец с кодом клиента. Убедитесь, что в файле есть столбец с названием, содержащим 'Код' и 'клиент'.")
            st.stop()
        
        # Сохраняем выбранные столбцы в session state
        st.session_state.year_month_col = year_month_col
        st.session_state.client_col = client_col
        
        # Построение матрицы
        if year_month_col and client_col:
            try:
                # Проверяем, есть ли уже вычисленные данные
                need_recompute = (
                    st.session_state.cohort_matrix is None or
                    st.session_state.sorted_periods is None or
                    st.session_state.year_month_col != year_month_col or
                    st.session_state.client_col != client_col
                )
                
                # Создаём контейнер для всего контента
                content_placeholder = st.empty()
                
                if need_recompute:
                    # Единый спиннер для всех расчётов - показываем только его
                    with content_placeholder.container():
                        with st.spinner("Расчёт и анализ данных..."):
                            # Построение когортной матрицы
                            cohort_matrix, sorted_periods = build_cohort_matrix(
                                df, 
                                year_month_col, 
                                client_col, 
                                value_type='clients'
                            )
                            st.session_state.cohort_matrix = cohort_matrix
                            st.session_state.sorted_periods = sorted_periods
                            
                            # Кэшируем множества клиентов по периодам для быстрого доступа в функциях получения клиентов
                            # Оптимизация: используем groupby вместо циклов с фильтрацией
                            period_clients_cache = {}
                            df_filtered = df[[year_month_col, client_col]].dropna()
                            for period, group in df_filtered.groupby(year_month_col):
                                period_clients_cache[period] = set(group[client_col].unique())
                            # Добавляем пустые множества для периодов без клиентов
                            for period in sorted_periods:
                                if period not in period_clients_cache:
                                    period_clients_cache[period] = set()
                            st.session_state.period_clients_cache = period_clients_cache
                            
                            # Кэшируем когорты клиентов (первый период появления каждого клиента)
                            client_cohorts_cache = get_client_cohorts(df, year_month_col, client_col, sorted_periods)
                            st.session_state.client_cohorts_cache = client_cohorts_cache
                            
                            # Вычисляем статистику по диагонали (количество клиентов в каждом периоде)
                            diagonal_values = {period: cohort_matrix.loc[period, period] for period in sorted_periods}
                            
                            # Находим максимум и минимум
                            max_clients = max(diagonal_values.values())
                            min_clients = min(diagonal_values.values())
                            max_period = [period for period, val in diagonal_values.items() if val == max_clients][0]
                            min_period = [period for period, val in diagonal_values.items() if val == min_clients][0]
                            
                            # Первый и последний период
                            first_period = sorted_periods[0]
                            last_period = sorted_periods[-1]
                            
                            # Сохраняем информацию в session state для отображения в правой колонке
                            st.session_state.cohort_info = {
                                'num_periods': len(sorted_periods),
                                'first_period': first_period,
                                'last_period': last_period,
                                'max_clients': max_clients,
                                'max_period': max_period,
                                'min_clients': min_clients,
                                'min_period': min_period
                            }
                            
                            # Построение всех остальных матриц внутри спиннера
                            st.session_state.accumulation_matrix = build_accumulation_matrix(df, year_month_col, client_col, sorted_periods)
                            st.session_state.accumulation_percent_matrix = build_accumulation_percent_matrix(st.session_state.accumulation_matrix, cohort_matrix)
                            st.session_state.inflow_matrix = build_inflow_matrix(st.session_state.accumulation_percent_matrix)
                            
                            # Кэшируем когорты клиентов (первый период появления каждого клиента)
                            client_cohorts_cache = get_client_cohorts(df, year_month_col, client_col, sorted_periods)
                            st.session_state.client_cohorts_cache = client_cohorts_cache
                            
                            st.session_state.churn_table = build_churn_table(df, year_month_col, client_col, sorted_periods, cohort_matrix, st.session_state.accumulation_matrix, st.session_state.accumulation_percent_matrix, client_cohorts_cache, period_clients_cache)
                            
                            # Кэшируем множества клиентов по периодам для быстрого доступа в функциях получения клиентов
                            # Оптимизация: используем groupby вместо циклов с фильтрацией
                            period_clients_cache = {}
                            df_filtered = df[[year_month_col, client_col]].dropna()
                            for period, group in df_filtered.groupby(year_month_col):
                                period_clients_cache[period] = set(group[client_col].unique())
                            # Добавляем пустые множества для периодов без клиентов
                            for period in sorted_periods:
                                if period not in period_clients_cache:
                                    period_clients_cache[period] = set()
                            st.session_state.period_clients_cache = period_clients_cache
                    
                    # После завершения всех расчётов очищаем placeholder и отображаем весь контент
                    content_placeholder.empty()
                else:
                    # Используем сохраненные данные
                    cohort_matrix = st.session_state.cohort_matrix
                    sorted_periods = st.session_state.sorted_periods
                    # Проверяем наличие остальных матриц
                    if st.session_state.get('accumulation_matrix') is None:
                        st.session_state.accumulation_matrix = build_accumulation_matrix(df, year_month_col, client_col, sorted_periods)
                    if st.session_state.get('accumulation_percent_matrix') is None:
                        st.session_state.accumulation_percent_matrix = build_accumulation_percent_matrix(st.session_state.accumulation_matrix, cohort_matrix)
                    if st.session_state.get('inflow_matrix') is None:
                        st.session_state.inflow_matrix = build_inflow_matrix(st.session_state.accumulation_percent_matrix)
                    if st.session_state.get('churn_table') is None:
                        client_cohorts_cache = st.session_state.get('client_cohorts_cache')
                        if client_cohorts_cache is None:
                            client_cohorts_cache = get_client_cohorts(df, year_month_col, client_col, sorted_periods)
                            st.session_state.client_cohorts_cache = client_cohorts_cache
                        period_clients_cache = st.session_state.get('period_clients_cache')
                        st.session_state.churn_table = build_churn_table(df, year_month_col, client_col, sorted_periods, cohort_matrix, st.session_state.accumulation_matrix, st.session_state.accumulation_percent_matrix, client_cohorts_cache, period_clients_cache)
                    
                    # Создаем кэш множеств клиентов, если его еще нет
                    if st.session_state.get('period_clients_cache') is None:
                        # Оптимизация: используем groupby вместо циклов с фильтрацией
                        period_clients_cache = {}
                        df_filtered = df[[year_month_col, client_col]].dropna()
                        for period, group in df_filtered.groupby(year_month_col):
                            period_clients_cache[period] = set(group[client_col].unique())
                        # Добавляем пустые множества для периодов без клиентов
                        for period in sorted_periods:
                            if period not in period_clients_cache:
                                period_clients_cache[period] = set()
                        st.session_state.period_clients_cache = period_clients_cache
                    
                    # Создаем кэш когорт клиентов, если его еще нет
                    if st.session_state.get('client_cohorts_cache') is None:
                        client_cohorts_cache = get_client_cohorts(df, year_month_col, client_col, sorted_periods)
                        st.session_state.client_cohorts_cache = client_cohorts_cache
                
                # Получаем информацию из session state
                info = st.session_state.cohort_info
                
                # Отображаем кнопки скачивания под блоком загрузки (горизонтально)
                if info:
                        # Создаем функцию для генерации полного отчёта
                        def create_full_report_excel():
                            """Создает полный Excel отчёт со всеми таблицами"""
                            buffer = io.BytesIO()
                            
                            # Импорты для форматирования Excel
                            from openpyxl.styles import Font
                            from openpyxl.utils import get_column_letter
                            
                            # Получаем данные из session state
                            cohort_matrix = st.session_state.cohort_matrix
                            sorted_periods = st.session_state.sorted_periods
                            df = st.session_state.df
                            year_month_col = st.session_state.year_month_col
                            client_col = st.session_state.client_col
                        
                            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                                workbook = writer.book
                                
                                # Получаем все матрицы
                                accumulation_matrix = build_accumulation_matrix(df, year_month_col, client_col, sorted_periods)
                                accumulation_percent_matrix = build_accumulation_percent_matrix(accumulation_matrix, cohort_matrix)
                                inflow_matrix = build_inflow_matrix(accumulation_percent_matrix)
                                
                                # Таблица 1: Динамика уникальных клиентов когорт
                                cohort_matrix_copy = cohort_matrix.copy()
                                cohort_matrix_copy.index.name = 'Когорта / Период'
                                cohort_matrix_copy.to_excel(writer, sheet_name="1. Динамика уникальных клиентов", startrow=0, index=True)
                                worksheet1 = writer.sheets["1. Динамика уникальных клиентов"]
                                # Используем специальное форматирование с горизонтальной динамикой
                                apply_excel_cohort_formatting(worksheet1, cohort_matrix.astype(float), sorted_periods)
                                
                                # Таблица 2: Динамика накопления возврата
                                accumulation_matrix_copy = accumulation_matrix.copy()
                                accumulation_matrix_copy.index.name = 'Когорта / Период'
                                accumulation_matrix_copy.to_excel(writer, sheet_name="2. Динамика накопления", startrow=0, index=True)
                                worksheet2 = writer.sheets["2. Динамика накопления"]
                                # Применяем форматирование со скрытием нулевых значений
                                apply_excel_color_formatting(worksheet2, accumulation_matrix.astype(float), hide_zeros=True)
                                # Форматируем значения как целые числа (только для непустых ячеек)
                                for row_idx in range(2, len(accumulation_matrix.index) + 2):
                                    for col_idx in range(2, len(accumulation_matrix.columns) + 2):
                                        cell = worksheet2.cell(row=row_idx, column=col_idx)
                                        if cell.value is not None and not isinstance(cell.value, str) and cell.value != "":
                                            cell.number_format = '0'  # Формат целого числа
                                
                                # Таблица 3: Динамика накопления возврата в %
                                accumulation_percent_matrix_copy = accumulation_percent_matrix.copy()
                                accumulation_percent_matrix_copy.index.name = 'Когорта / Период'
                                accumulation_percent_matrix_copy.to_excel(writer, sheet_name="3. Динамика накопления %", startrow=0, index=True)
                                worksheet3 = writer.sheets["3. Динамика накопления %"]
                                # Используем специальное форматирование для процентов
                                apply_excel_percent_formatting(worksheet3, accumulation_percent_matrix, sorted_periods)
                                
                                # Таблица 4: Приток возврата в %
                                inflow_matrix_copy = inflow_matrix.copy()
                                inflow_matrix_copy.index.name = 'Когорта / Период'
                                inflow_matrix_copy.to_excel(writer, sheet_name="4. Приток возврата %", startrow=0, index=True)
                                worksheet4 = writer.sheets["4. Приток возврата %"]
                                # Используем специальное форматирование для процентов притока
                                apply_excel_inflow_formatting(worksheet4, inflow_matrix, sorted_periods)
                                
                                # Таблица 5: Отток клиентов из категории
                                churn_table_full = build_churn_table(df, year_month_col, client_col, sorted_periods, cohort_matrix, accumulation_matrix, accumulation_percent_matrix, None, None)
                                churn_table_copy = churn_table_full.copy()
                                # Не конвертируем проценты в строки - сохраняем как числа для возможности расчетов
                                churn_table_copy.to_excel(writer, sheet_name="5. Отток клиентов из категории", startrow=0, index=False)
                                worksheet5 = writer.sheets["5. Отток клиентов из категории"]
                                # Форматируем значения: числа как целые, проценты как проценты
                                from openpyxl.styles import Alignment as ExcelAlignment
                                for row_idx in range(2, len(churn_table_copy) + 2):
                                    for col_idx in range(1, len(churn_table_copy.columns) + 1):
                                        cell = worksheet5.cell(row=row_idx, column=col_idx)
                                        cell.alignment = ExcelAlignment(horizontal="center", vertical="center")
                                        col_name = churn_table_copy.columns[col_idx - 1]
                                        if col_name in ['Кол-во клиентов когорты', 'Накопительное кол-во возврата', 'Отток кол-во']:
                                            # Колонки с числами
                                            if cell.value is not None and not isinstance(cell.value, str):
                                                cell.number_format = '0'  # Формат целого числа
                                        elif col_name in ['Накопительный % возврата', 'Отток %']:
                                            # Колонки с процентами - сохраняем как число (уже в процентах, конвертируем в долю)
                                            if cell.value is not None and not isinstance(cell.value, str):
                                                # Значение уже в процентах (например, 45.7), конвертируем в долю (0.457)
                                                cell.value = float(cell.value) / 100.0
                                                cell.number_format = '0.0%'  # Процентный формат Excel
                                
                                # Таблица 7: Присутствие клиентов оттока когорты в других категориях товаров
                                if 'df_categories' in st.session_state and st.session_state.df_categories is not None and \
                                   'categories_list' in st.session_state and st.session_state.categories_list is not None and \
                                   'group_col_name' in st.session_state and st.session_state.group_col_name is not None and \
                                   'year_month_col_name' in st.session_state and 'client_code_col_name' in st.session_state:
                                    
                                    df_categories = st.session_state.df_categories
                                    categories = st.session_state.categories_list
                                    group_col = st.session_state.group_col_name
                                    year_month_col_cat = st.session_state.get('year_month_col_name', None)
                                    client_code_col = st.session_state.get('client_code_col_name', None)
                                    
                                    start_row_cohorts = 0
                                    worksheet_cohorts = None
                                    
                                    # Добавляем данные с листа 6 (если есть)
                                    if 'category_summary_table' in st.session_state and st.session_state.category_summary_table is not None:
                                        summary_table_excel = st.session_state.category_summary_table.copy()
                                        summary_table_excel.index.name = 'Метрика / Когорта'
                                        summary_table_excel.to_excel(writer, sheet_name="7. Присутствие когорты в других категориях", startrow=start_row_cohorts, index=True)
                                        worksheet_cohorts = writer.sheets["7. Присутствие когорты в других категориях"]
                                        
                                        # Форматируем верхнюю таблицу
                                        for row_idx in range(start_row_cohorts + 2, start_row_cohorts + len(summary_table_excel.index) + 2):
                                            for col_idx in range(2, len(summary_table_excel.columns) + 2):
                                                cell = worksheet_cohorts.cell(row=row_idx, column=col_idx)
                                                cell.alignment = ExcelAlignment(horizontal="center", vertical="center")
                                                row_name = summary_table_excel.index[row_idx - start_row_cohorts - 2]
                                                
                                                if cell.value is not None and not isinstance(cell.value, str):
                                                    if row_name == 'Доля оттока из сети от когорты':
                                                        cell.value = float(cell.value) / 100.0
                                                        cell.number_format = '0.0%'
                                                    else:
                                                        cell.number_format = '0'
                                        
                                        # Форматируем заголовок строки
                                        for row_idx in range(start_row_cohorts + 2, start_row_cohorts + len(summary_table_excel.index) + 2):
                                            cell = worksheet_cohorts.cell(row=row_idx, column=1)
                                            cell.alignment = ExcelAlignment(horizontal="left", vertical="center")
                                        
                                        start_row_cohorts = start_row_cohorts + len(summary_table_excel.index) + 3
                                    
                                    if 'category_cohort_table' in st.session_state and st.session_state.category_cohort_table is not None:
                                        category_table_excel = st.session_state.category_cohort_table.copy()
                                        category_table_excel.index.name = 'Категория / Когорта'
                                        
                                        if worksheet_cohorts is None:
                                            category_table_excel.to_excel(writer, sheet_name="7. Присутствие когорты в других категориях", startrow=start_row_cohorts, index=True)
                                            worksheet_cohorts = writer.sheets["7. Присутствие когорты в других категориях"]
                                        else:
                                            category_table_excel.to_excel(writer, sheet_name="7. Присутствие когорты в других категориях", startrow=start_row_cohorts, index=True)
                                        
                                        # Форматируем таблицу с категориями
                                        for row_idx in range(start_row_cohorts + 2, start_row_cohorts + len(category_table_excel.index) + 2):
                                            for col_idx in range(2, len(category_table_excel.columns) + 2):
                                                cell = worksheet_cohorts.cell(row=row_idx, column=col_idx)
                                                cell.alignment = ExcelAlignment(horizontal="center", vertical="center")
                                                if cell.value is not None and not isinstance(cell.value, str):
                                                    cell.number_format = '0'
                                        
                                        # Форматируем заголовок строки
                                        for row_idx in range(start_row_cohorts + 2, start_row_cohorts + len(category_table_excel.index) + 2):
                                            cell = worksheet_cohorts.cell(row=row_idx, column=1)
                                            cell.alignment = ExcelAlignment(horizontal="left", vertical="center")
                                        
                                        start_row_cohorts = start_row_cohorts + len(category_table_excel.index) + 3
                                    
                                    # Для каждой когорты создаем таблицу
                                    for cohort_idx, selected_cohort in enumerate(sorted_periods):
                                        # Определяем периоды начиная с выбранной когорты
                                        cohort_index = sorted_periods.index(selected_cohort) if selected_cohort in sorted_periods else 0
                                        periods_from_cohort = sorted_periods[cohort_index:]
                                        
                                        # Получаем клиентов оттока для выбранной когорты
                                        period_clients_cache = st.session_state.get('period_clients_cache', None)
                                        client_cohorts_cache = st.session_state.get('client_cohorts_cache', None)
                                        churn_clients_set = set(get_churn_clients(df, year_month_col, client_col, sorted_periods, selected_cohort, period_clients_cache, client_cohorts_cache))
                                        churn_clients_set = {str(client) for client in churn_clients_set}
                                        
                                        # Создаем таблицу: категории по строкам, периоды по столбцам
                                        category_period_table = pd.DataFrame(index=categories, columns=periods_from_cohort)
                                        
                                        # Словари для итогов
                                        period_unique_clients = {period: set() for period in periods_from_cohort}
                                        category_unique_clients = {category: set() for category in categories}
                                        
                                        # Если есть столбец "Год-месяц", используем его для фильтрации по периодам
                                        if year_month_col_cat is not None:
                                            for period in periods_from_cohort:
                                                period_data = df_categories[df_categories[year_month_col_cat] == period]
                                                
                                                for category in categories:
                                                    category_period_data = period_data[
                                                        (period_data[group_col] == category) & 
                                                        (period_data[client_code_col].notna())
                                                    ]
                                                    
                                                    category_period_clients = set(
                                                        category_period_data[client_code_col].dropna().astype(str).unique()
                                                    )
                                                    
                                                    intersection = churn_clients_set & category_period_clients
                                                    category_period_table.loc[category, period] = len(intersection)
                                                    
                                                    period_unique_clients[period].update(intersection)
                                                    category_unique_clients[category].update(intersection)
                                        else:
                                            category_clients_dict = {}
                                            for category in categories:
                                                category_data = df_categories[df_categories[group_col] == category]
                                                client_codes = set(category_data[client_code_col].dropna().astype(str).unique())
                                                category_clients_dict[category] = client_codes
                                            
                                            for period in periods_from_cohort:
                                                for category in categories:
                                                    category_clients_set = category_clients_dict.get(category, set())
                                                    intersection = churn_clients_set & category_clients_set
                                                    category_period_table.loc[category, period] = len(intersection)
                                                    
                                                    period_unique_clients[period].update(intersection)
                                                    category_unique_clients[category].update(intersection)
                                        
                                        # Заполняем NaN нулями
                                        category_period_table = category_period_table.fillna(0).astype(int)
                                        
                                        # Создаем итоговую строку и столбец
                                        totals_row = pd.Series(
                                            {period: len(period_unique_clients[period]) for period in periods_from_cohort},
                                            name='Итого клиентов'
                                        )
                                        
                                        totals_col = pd.Series(
                                            {category: len(category_unique_clients[category]) for category in categories},
                                            name='Итого'
                                        )
                                        
                                        # Добавляем итоги в таблицу
                                        category_period_table_with_totals = category_period_table.copy()
                                        category_period_table_with_totals.loc['Итого клиентов'] = totals_row
                                        category_period_table_with_totals['Итого'] = totals_col
                                        
                                        # Вычисляем значение для ячейки пересечения
                                        all_category_clients = set()
                                        if year_month_col_cat is not None:
                                            for category in categories:
                                                category_data = df_categories[df_categories[group_col] == category]
                                                category_data_filtered = category_data[category_data[year_month_col_cat].isin(periods_from_cohort)]
                                                category_clients = set(category_data_filtered[client_code_col].dropna().astype(str).unique())
                                                all_category_clients.update(category_clients)
                                        else:
                                            for category in categories:
                                                category_data = df_categories[df_categories[group_col] == category]
                                                category_clients = set(category_data[client_code_col].dropna().astype(str).unique())
                                                all_category_clients.update(category_clients)
                                        
                                        present_in_categories = churn_clients_set & all_category_clients
                                        category_period_table_with_totals.loc['Итого клиентов', 'Итого'] = len(present_in_categories)
                                        
                                        # Переупорядочиваем строки и столбцы
                                        new_index = ['Итого клиентов'] + [cat for cat in categories]
                                        category_period_table_with_totals = category_period_table_with_totals.reindex(new_index)
                                        new_columns = ['Итого'] + list(periods_from_cohort)
                                        category_period_table_with_totals = category_period_table_with_totals[new_columns]
                                        
                                        # Добавляем заголовок когорты
                                        if worksheet_cohorts is None:
                                            # Создаем новый лист
                                            category_period_table_with_totals.to_excel(
                                                writer, 
                                                sheet_name="7. Присутствие когорты в других категориях", 
                                                startrow=start_row_cohorts, 
                                                index=True
                                            )
                                            worksheet_cohorts = writer.sheets["7. Присутствие когорты в других категориях"]
                                            # Добавляем заголовок когорты
                                            last_col_letter = get_column_letter(len(new_columns) + 1)
                                            worksheet_cohorts.cell(row=start_row_cohorts + 1, column=1, value=f"Когорта: {selected_cohort}")
                                            worksheet_cohorts.merge_cells(f'A{start_row_cohorts + 1}:{last_col_letter}{start_row_cohorts + 1}')
                                            header_cell = worksheet_cohorts.cell(row=start_row_cohorts + 1, column=1)
                                            header_cell.font = Font(bold=True, size=12)
                                            header_cell.alignment = ExcelAlignment(horizontal="center", vertical="center")
                                            start_row_cohorts += 2
                                        else:
                                            # Добавляем заголовок когорты
                                            last_col_letter = get_column_letter(len(new_columns) + 1)
                                            worksheet_cohorts.cell(row=start_row_cohorts + 1, column=1, value=f"Когорта: {selected_cohort}")
                                            worksheet_cohorts.merge_cells(f'A{start_row_cohorts + 1}:{last_col_letter}{start_row_cohorts + 1}')
                                            header_cell = worksheet_cohorts.cell(row=start_row_cohorts + 1, column=1)
                                            header_cell.font = Font(bold=True, size=12)
                                            header_cell.alignment = ExcelAlignment(horizontal="center", vertical="center")
                                            start_row_cohorts += 2
                                            
                                            # Записываем таблицу на тот же лист
                                            category_period_table_with_totals.to_excel(
                                                writer, 
                                                sheet_name="7. Присутствие когорты в других категориях", 
                                                startrow=start_row_cohorts, 
                                                index=True
                                            )
                                        
                                        # Форматируем таблицу
                                        for row_idx in range(start_row_cohorts + 2, start_row_cohorts + len(category_period_table_with_totals.index) + 2):
                                            for col_idx in range(2, len(category_period_table_with_totals.columns) + 2):
                                                cell = worksheet_cohorts.cell(row=row_idx, column=col_idx)
                                                cell.alignment = ExcelAlignment(horizontal="center", vertical="center")
                                                if cell.value is not None and not isinstance(cell.value, str):
                                                    cell.number_format = '0'
                                        
                                        # Форматируем заголовок строки
                                        for row_idx in range(start_row_cohorts + 2, start_row_cohorts + len(category_period_table_with_totals.index) + 2):
                                            cell = worksheet_cohorts.cell(row=row_idx, column=1)
                                            cell.alignment = ExcelAlignment(horizontal="left", vertical="center")
                                        
                                        # Обновляем начальную строку для следующей таблицы (таблица + 2 пустые строки)
                                        start_row_cohorts = start_row_cohorts + len(category_period_table_with_totals.index) + 3
                                
                                # Таблица 8: Сводная таблица по всем когортам
                                if st.session_state.get('churn_table') is not None:
                                    churn_table = st.session_state.churn_table
                                    
                                    # Создаем сводную таблицу (та же логика, что и в интерфейсе)
                                    summary_data = {}
                                    
                                    # 1. Кол-во клиентов в когорте
                                    summary_data['Кол-во клиентов в когорте'] = {}
                                    for _, row in churn_table.iterrows():
                                        cohort = row['Когорта']
                                        summary_data['Кол-во клиентов в когорте'][cohort] = int(row['Кол-во клиентов когорты'])
                                    
                                    # 2. Накопительное кол-во вернувшихся в категорию
                                    summary_data['Накопительное кол-во вернувшихся в категорию'] = {}
                                    for _, row in churn_table.iterrows():
                                        cohort = row['Когорта']
                                        summary_data['Накопительное кол-во вернувшихся в категорию'][cohort] = int(row['Накопительное кол-во возврата'])
                                    
                                    # 3. Накопительное кол-во вернувшихся в категорию %
                                    summary_data['Накопительное кол-во вернувшихся в категорию %'] = {}
                                    for _, row in churn_table.iterrows():
                                        cohort = row['Когорта']
                                        summary_data['Накопительное кол-во вернувшихся в категорию %'][cohort] = row['Накопительный % возврата']
                                    
                                    # 4. Отток из категории когорты
                                    summary_data['Отток из категории когорты'] = {}
                                    for _, row in churn_table.iterrows():
                                        cohort = row['Когорта']
                                        summary_data['Отток из категории когорты'][cohort] = int(row['Отток кол-во'])
                                    
                                    # 5. Отток из категории когорты %
                                    summary_data['Отток из категории когорты %'] = {}
                                    for _, row in churn_table.iterrows():
                                        cohort = row['Когорта']
                                        summary_data['Отток из категории когорты %'][cohort] = row['Отток %']
                                    
                                    # Инициализируем словари для метрик 6-9
                                    summary_data['Кол-во клиентов когорты в других категориях после месяца когорты'] = {}
                                    summary_data['Кол-во клиентов когорты в других категориях после месяца когорты %'] = {}
                                    summary_data['Отток из сети'] = {}
                                    summary_data['Отток из сети %'] = {}
                                    
                                    for cohort in sorted_periods:
                                        summary_data['Кол-во клиентов когорты в других категориях после месяца когорты'][cohort] = 0
                                        summary_data['Кол-во клиентов когорты в других категориях после месяца когорты %'][cohort] = 0.0
                                        summary_data['Отток из сети'][cohort] = 0
                                        summary_data['Отток из сети %'][cohort] = 0.0
                                    
                                    # 6-9. Данные о присутствии в других категориях после месяца когорты и оттоке из сети
                                    if 'category_summary_table' in st.session_state and st.session_state.category_summary_table is not None:
                                        category_summary = st.session_state.category_summary_table
                                        
                                        if 'Итого присутствуют в других категориях после месяца когорты' in category_summary.index:
                                            for cohort in sorted_periods:
                                                if cohort in category_summary.columns:
                                                    value = category_summary.loc['Итого присутствуют в других категориях после месяца когорты', cohort]
                                                    summary_data['Кол-во клиентов когорты в других категориях после месяца когорты'][cohort] = int(value) if pd.notna(value) else 0
                                        
                                        if 'Доля присутствуют в других категориях после месяца когорты' in category_summary.index:
                                            for cohort in sorted_periods:
                                                if cohort in category_summary.columns:
                                                    value = category_summary.loc['Доля присутствуют в других категориях после месяца когорты', cohort]
                                                    if pd.notna(value):
                                                        summary_data['Кол-во клиентов когорты в других категориях после месяца когорты %'][cohort] = value
                                        else:
                                            # Вычисляем процент вручную
                                            for cohort in sorted_periods:
                                                cohort_size = summary_data['Кол-во клиентов в когорте'].get(cohort, 0)
                                                present_after_count = summary_data['Кол-во клиентов когорты в других категориях после месяца когорты'].get(cohort, 0)
                                                if cohort_size > 0:
                                                    percent = (present_after_count / cohort_size) * 100
                                                    summary_data['Кол-во клиентов когорты в других категориях после месяца когорты %'][cohort] = percent
                                        
                                        if 'Отток из сети' in category_summary.index:
                                            for cohort in sorted_periods:
                                                if cohort in category_summary.columns:
                                                    value = category_summary.loc['Отток из сети', cohort]
                                                    summary_data['Отток из сети'][cohort] = int(value) if pd.notna(value) else 0
                                        
                                        if 'Доля оттока из сети от когорты' in category_summary.index:
                                            for cohort in sorted_periods:
                                                if cohort in category_summary.columns:
                                                    value = category_summary.loc['Доля оттока из сети от когорты', cohort]
                                                    if pd.notna(value):
                                                        summary_data['Отток из сети %'][cohort] = value
                                        
                                    # Создаем DataFrame
                                    summary_df = pd.DataFrame(summary_data, index=sorted_periods).T
                                    summary_df.index.name = 'Метрика / Когорта'
                                    
                                    # Записываем в Excel
                                    summary_df.to_excel(writer, sheet_name="8. Сводная таблица по всем когортам", startrow=0, index=True)
                                    worksheet_summary = writer.sheets["8. Сводная таблица по всем когортам"]
                                    
                                    # Форматируем таблицу
                                    for row_idx in range(2, len(summary_df.index) + 2):
                                        for col_idx in range(2, len(summary_df.columns) + 2):
                                            cell = worksheet_summary.cell(row=row_idx, column=col_idx)
                                            cell.alignment = ExcelAlignment(horizontal="center", vertical="center")
                                            row_name = summary_df.index[row_idx - 2]
                                            
                                            if cell.value is not None and not isinstance(cell.value, str):
                                                if '%' in row_name:
                                                    # Процентные колонки
                                                    cell.value = float(cell.value) / 100.0 if isinstance(cell.value, (int, float)) and cell.value > 1 else float(cell.value)
                                                    cell.number_format = '0.0%'
                                                else:
                                                    # Числовые колонки
                                                    cell.number_format = '0'
                                    
                                    # Форматируем заголовок строки
                                    for row_idx in range(2, len(summary_df.index) + 2):
                                        cell = worksheet_summary.cell(row=row_idx, column=1)
                                        cell.alignment = ExcelAlignment(horizontal="left", vertical="center")
                                
                                # Удаляем пустой лист по умолчанию
                                if 'Sheet' in workbook.sheetnames:
                                    workbook.remove(workbook['Sheet'])
                            
                            buffer.seek(0)
                            return buffer.getvalue()
                        
                        # CSS для увеличения размера кнопок загрузки
                        st.markdown("""
                        <style>
                        div[data-testid="stDownloadButton"] > button {
                            height: 60px !important;
                            font-size: 20px !important;
                            font-weight: bold !important;
                            padding: 15px 30px !important;
                        }
                        div[data-testid="stDownloadButton"] > button > div > p {
                            font-size: 20px !important;
                            font-weight: bold !important;
                        }
                        </style>
                        """, unsafe_allow_html=True)
                        
                        # Создаем колонки для горизонтального размещения кнопок
                        col_excel_button, col_pdf_button = st.columns(2)
                        
                        # Генерируем файл каждый раз при рендеринге (данные могут обновиться)
                        # Всегда генерируем отчет заново, чтобы включить все актуальные данные
                        try:
                            excel_data_full = create_full_report_excel()
                            # Сохраняем для возможного использования в будущем
                            st.session_state.excel_report_data = excel_data_full
                        except Exception as e:
                            # Если ошибка, используем сохраненный файл как fallback
                            if 'excel_report_data' in st.session_state and st.session_state.excel_report_data is not None:
                                excel_data_full = st.session_state.excel_report_data
                                st.warning(f"Использован сохраненный отчет. Ошибка при генерации: {str(e)}")
                            else:
                                st.error(f"Ошибка при генерации отчета: {str(e)}")
                                excel_data_full = b""  # Пустой файл
                        
                        with col_excel_button:
                            st.download_button(
                                label="📥 Скачать полный отчёт в Excel",
                                data=excel_data_full,
                                file_name=f"полный_отчёт_когортный_анализ_{info['first_period']}_{info['last_period']}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True,
                                key="download_full_report"
                            )
                        
                        # Создаем функцию для генерации аналитического PDF отчёта
                        def create_analysis_pdf():
                            """Создает PDF отчёт с графиками и анализом"""
                            buffer = io.BytesIO()
                            
                            # Регистрируем шрифт с поддержкой кириллицы
                            font_name = 'Helvetica'
                            font_name_bold = 'Helvetica-Bold'
                            
                            try:
                                # Пытаемся найти системный шрифт с поддержкой кириллицы
                                if platform.system() == 'Windows':
                                    # Пути к стандартным шрифтам Windows с поддержкой кириллицы
                                    windows_fonts = [
                                        r'C:\Windows\Fonts\arial.ttf',
                                        r'C:\Windows\Fonts\calibri.ttf',
                                        r'C:\Windows\Fonts\comic.ttf',
                                        r'C:\Windows\Fonts\cour.ttf',
                                    ]
                                    
                                    # Регистрируем первый доступный шрифт
                                    for font_path in windows_fonts:
                                        if os.path.exists(font_path):
                                            try:
                                                font_name = 'CyrillicFont'
                                                font_name_bold = 'CyrillicFont-Bold'
                                                pdfmetrics.registerFont(TTFont(font_name, font_path))
                                                pdfmetrics.registerFont(TTFont(font_name_bold, font_path))
                                                break
                                            except Exception as e:
                                                continue
                                elif platform.system() == 'Linux':
                                    # Пути к стандартным шрифтам Linux с поддержкой кириллицы
                                    linux_fonts = [
                                        '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
                                        '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf',
                                        '/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf',
                                        '/usr/share/fonts/truetype/ttf-dejavu/DejaVuSans.ttf',
                                        '/usr/share/fonts/TTF/DejaVuSans.ttf',
                                    ]
                                    
                                    # Регистрируем первый доступный шрифт
                                    for font_path in linux_fonts:
                                        if os.path.exists(font_path):
                                            try:
                                                font_name = 'CyrillicFont'
                                                font_name_bold = 'CyrillicFont-Bold'
                                                pdfmetrics.registerFont(TTFont(font_name, font_path))
                                                pdfmetrics.registerFont(TTFont(font_name_bold, font_path))
                                                break
                                            except Exception as e:
                                                continue
                            except Exception as e:
                                pass  # Используем стандартные шрифты в случае ошибки
                            
                            # Получаем данные из session state
                            cohort_matrix = st.session_state.cohort_matrix
                            sorted_periods = st.session_state.sorted_periods
                            accumulation_matrix = st.session_state.accumulation_matrix
                            accumulation_percent_matrix = st.session_state.accumulation_percent_matrix
                            inflow_matrix = st.session_state.inflow_matrix
                            churn_table = st.session_state.churn_table
                            
                            # Создаем PDF документ
                            doc = SimpleDocTemplate(buffer, pagesize=A4)
                            story = []
                            styles = getSampleStyleSheet()
                            
                            # Стили с поддержкой кириллицы
                            title_style = ParagraphStyle(
                                'CustomTitle',
                                parent=styles['Heading1'],
                                fontName=font_name_bold,
                                fontSize=24,
                                textColor=colors.HexColor('#1f77b4'),
                                spaceAfter=30,
                                alignment=TA_CENTER
                            )
                            
                            heading_style = ParagraphStyle(
                                'CustomHeading',
                                parent=styles['Heading2'],
                                fontName=font_name_bold,
                                fontSize=16,
                                textColor=colors.HexColor('#1f77b4'),
                                spaceAfter=12,
                                spaceBefore=12
                            )
                            
                            # Стиль для обычного текста с поддержкой кириллицы
                            normal_style = ParagraphStyle(
                                'CustomNormal',
                                parent=styles['Normal'],
                                fontName=font_name,
                                fontSize=10
                            )
                            
                            # Стиль для заголовков третьего уровня с поддержкой кириллицы
                            heading3_style = ParagraphStyle(
                                'CustomHeading3',
                                parent=styles['Heading3'],
                                fontName=font_name_bold,
                                fontSize=12,
                                textColor=colors.HexColor('#1f77b4'),
                                spaceAfter=8,
                                spaceBefore=8
                            )
                            
                            # Титульная страница
                            story.append(Paragraph("КОГОРТНЫЙ АНАЛИЗ", title_style))
                            story.append(Spacer(1, 0.3*inch))
                            story.append(Paragraph(f"Период анализа: {info['first_period']} - {info['last_period']}", normal_style))
                            story.append(Paragraph(f"Количество когорт: {info['num_periods']}", normal_style))
                            story.append(Paragraph(f"Дата формирования: {datetime.now().strftime('%d.%m.%Y %H:%M')}", normal_style))
                            story.append(PageBreak())
                            
                            # Раздел 1: Общая статистика
                            story.append(Paragraph("1. ОБЩАЯ СТАТИСТИКА", heading_style))
                            
                            # Диагональные значения (размер когорт)
                            diagonal_values = {period: cohort_matrix.loc[period, period] for period in sorted_periods}
                            
                            stats_data = [
                                ['Метрика', 'Значение'],
                                ['Всего когорт', str(info['num_periods'])],
                                ['Период начала', info['first_period']],
                                ['Период окончания', info['last_period']],
                                ['Максимальный размер когорты', f"{int(info['max_clients'])} ({info['max_period']})"],
                                ['Минимальный размер когорты', f"{int(info['min_clients'])} ({info['min_period']})"],
                                ['Средний размер когорты', f"{int(np.mean(list(diagonal_values.values())))}"],
                                ['Общее количество уникальных клиентов', f"{int(sum(diagonal_values.values()))}"]
                            ]
                            
                            stats_table = Table(stats_data, colWidths=[4*inch, 3*inch])
                            stats_table.setStyle(TableStyle([
                                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f77b4')),
                                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                                ('FONTNAME', (0, 0), (-1, 0), font_name_bold),
                                ('FONTNAME', (0, 1), (-1, -1), font_name),
                                ('FONTSIZE', (0, 0), (-1, 0), 12),
                                ('FONTSIZE', (0, 1), (-1, -1), 10),
                                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                                ('GRID', (0, 0), (-1, -1), 1, colors.black)
                            ]))
                            story.append(stats_table)
                            story.append(Spacer(1, 0.3*inch))
                            
                            # График 1: Динамика размера когорт
                            story.append(Paragraph("2. ДИНАМИКА РАЗМЕРА КОГОРТ", heading_style))
                            
                            fig, ax = plt.subplots(figsize=(10, 6))
                            cohort_sizes = [diagonal_values[p] for p in sorted_periods]
                            ax.plot(range(len(sorted_periods)), cohort_sizes, marker='o', linewidth=2, markersize=8, color='#1f77b4')
                            ax.set_xlabel('Период', fontsize=12, fontweight='bold')
                            ax.set_ylabel('Количество клиентов', fontsize=12, fontweight='bold')
                            ax.set_title('Динамика размера когорт по периодам', fontsize=14, fontweight='bold', pad=20)
                            ax.set_xticks(range(len(sorted_periods)))
                            ax.set_xticklabels(sorted_periods, rotation=45, ha='right')
                            ax.grid(True, alpha=0.3)
                            ax.set_facecolor('#f8f9fa')
                            
                            for i, (period, size) in enumerate(zip(sorted_periods, cohort_sizes)):
                                ax.annotate(f'{int(size)}', (i, size), textcoords="offset points", xytext=(0,10), ha='center', fontsize=9)
                            
                            plt.tight_layout()
                            img_buffer1 = io.BytesIO()
                            plt.savefig(img_buffer1, format='png', dpi=150, bbox_inches='tight')
                            img_buffer1.seek(0)
                            plt.close()
                            
                            img1 = Image(img_buffer1, width=6*inch, height=3.6*inch)
                            story.append(img1)
                            story.append(Spacer(1, 0.3*inch))
                            
                            # График 2: Тепловая карта возврата в %
                            story.append(Paragraph("3. ТЕПЛОВАЯ КАРТА ВОЗВРАТА В %", heading_style))
                            
                            # Создаём упрощённую матрицу для визуализации (первые 15 когорт и периодов)
                            max_cohorts = min(15, len(sorted_periods))
                            matrix_vis = accumulation_percent_matrix.iloc[:max_cohorts, :max_cohorts]
                            
                            fig, ax = plt.subplots(figsize=(12, 10))
                            sns.heatmap(matrix_vis, annot=True, fmt='.1f', cmap='RdYlGn', 
                                       cbar_kws={'label': 'Процент возврата (%)'}, 
                                       ax=ax, vmin=0, vmax=100, linewidths=0.5, linecolor='gray')
                            ax.set_title('Тепловая карта накопления возврата клиентов (%)', fontsize=14, fontweight='bold', pad=20)
                            ax.set_xlabel('Период', fontsize=12, fontweight='bold')
                            ax.set_ylabel('Когорта', fontsize=12, fontweight='bold')
                            
                            plt.tight_layout()
                            img_buffer2 = io.BytesIO()
                            plt.savefig(img_buffer2, format='png', dpi=150, bbox_inches='tight')
                            img_buffer2.seek(0)
                            plt.close()
                            
                            img2 = Image(img_buffer2, width=6*inch, height=5*inch)
                            story.append(img2)
                            story.append(Spacer(1, 0.3*inch))
                            
                            # График 3: Отток по когортам
                            story.append(Paragraph("4. АНАЛИЗ ОТТОКА КЛИЕНТОВ", heading_style))
                            
                            fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
                            
                            # Столбчатая диаграмма оттока в количестве
                            churn_counts = churn_table['Отток кол-во'].values[:15]
                            cohorts_display = churn_table['Когорта'].values[:15]
                            
                            colors_churn = ['#d62728' if x > churn_table['Отток кол-во'].mean() else '#ff7f0e' for x in churn_counts]
                            ax1.barh(range(len(cohorts_display)), churn_counts, color=colors_churn)
                            ax1.set_yticks(range(len(cohorts_display)))
                            ax1.set_yticklabels(cohorts_display, fontsize=9)
                            ax1.set_xlabel('Количество клиентов оттока', fontsize=11, fontweight='bold')
                            ax1.set_title('Отток клиентов из категории по когортам', fontsize=12, fontweight='bold')
                            ax1.grid(True, alpha=0.3, axis='x')
                            
                            # Столбчатая диаграмма оттока в процентах
                            churn_percents = churn_table['Отток %'].values[:15]
                            colors_churn_pct = ['#d62728' if x > churn_table['Отток %'].mean() else '#ff7f0e' for x in churn_percents]
                            ax2.barh(range(len(cohorts_display)), churn_percents, color=colors_churn_pct)
                            ax2.set_yticks(range(len(cohorts_display)))
                            ax2.set_yticklabels(cohorts_display, fontsize=9)
                            ax2.set_xlabel('Процент оттока (%)', fontsize=11, fontweight='bold')
                            ax2.set_title('Процент оттока по когортам', fontsize=12, fontweight='bold')
                            ax2.grid(True, alpha=0.3, axis='x')
                            
                            plt.tight_layout()
                            img_buffer4 = io.BytesIO()
                            plt.savefig(img_buffer4, format='png', dpi=150, bbox_inches='tight')
                            img_buffer4.seek(0)
                            plt.close()
                            
                            img4 = Image(img_buffer4, width=7*inch, height=3.6*inch)
                            story.append(img4)
                            story.append(Spacer(1, 0.3*inch))
                            
                            # Таблицы с ключевыми метриками
                            story.append(Paragraph("5. КЛЮЧЕВЫЕ МЕТРИКИ", heading_style))
                            
                            # Топ-5 когорт по размеру
                            story.append(Paragraph("Топ-5 когорт по размеру:", heading3_style))
                            top5_size = sorted(diagonal_values.items(), key=lambda x: x[1], reverse=True)[:5]
                            top5_data = [['Место', 'Когорта', 'Количество клиентов']]
                            for i, (period, size) in enumerate(top5_size, 1):
                                top5_data.append([str(i), period, str(int(size))])
                            
                            top5_table = Table(top5_data, colWidths=[0.8*inch, 2.5*inch, 2*inch])
                            top5_table.setStyle(TableStyle([
                                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f77b4')),
                                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                                ('FONTNAME', (0, 0), (-1, 0), font_name_bold),
                                ('FONTNAME', (0, 1), (-1, -1), font_name),
                                ('FONTSIZE', (0, 0), (-1, 0), 10),
                                ('FONTSIZE', (0, 1), (-1, -1), 10),
                                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                                ('GRID', (0, 0), (-1, -1), 1, colors.black)
                            ]))
                            story.append(top5_table)
                            story.append(Spacer(1, 0.2*inch))
                            
                            # Топ-5 когорт по проценту возврата
                            story.append(Paragraph("Топ-5 когорт по проценту возврата:", heading3_style))
                            churn_sorted_return = churn_table.sort_values('Накопительный % возврата', ascending=False)
                            top5_return_data = [['Место', 'Когорта', 'Процент возврата', 'Размер когорты']]
                            for i, row in enumerate(churn_sorted_return.head(5).itertuples(index=False), 1):
                                top5_return_data.append([
                                    str(i), 
                                    row[0],  # Когорта
                                    f"{row[3]:.1f}%",  # Накопительный % возврата
                                    str(int(row[1]))  # Кол-во клиентов когорты
                                ])
                            
                            top5_return_table = Table(top5_return_data, colWidths=[0.8*inch, 2*inch, 1.5*inch, 1.5*inch])
                            top5_return_table.setStyle(TableStyle([
                                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2ca02c')),
                                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                                ('FONTNAME', (0, 0), (-1, 0), font_name_bold),
                                ('FONTNAME', (0, 1), (-1, -1), font_name),
                                ('FONTSIZE', (0, 0), (-1, 0), 10),
                                ('FONTSIZE', (0, 1), (-1, -1), 10),
                                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                                ('GRID', (0, 0), (-1, -1), 1, colors.black)
                            ]))
                            story.append(top5_return_table)
                            story.append(Spacer(1, 0.2*inch))
                            
                            # Когорты с максимальным оттоком
                            story.append(Paragraph("Топ-5 когорт с наибольшим оттоком:", heading3_style))
                            churn_sorted_churn = churn_table.sort_values('Отток %', ascending=False)
                            top5_churn_data = [['Место', 'Когорта', 'Отток (%)', 'Отток (кол-во)']]
                            for i, row in enumerate(churn_sorted_churn.head(5).itertuples(index=False), 1):
                                top5_churn_data.append([
                                    str(i),
                                    row[0],  # Когорта
                                    f"{row[5]:.1f}%",  # Отток %
                                    str(int(row[4]))  # Отток кол-во
                                ])
                            
                            top5_churn_table = Table(top5_churn_data, colWidths=[0.8*inch, 2*inch, 1.5*inch, 1.5*inch])
                            top5_churn_table.setStyle(TableStyle([
                                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#d62728')),
                                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                                ('FONTNAME', (0, 0), (-1, 0), font_name_bold),
                                ('FONTNAME', (0, 1), (-1, -1), font_name),
                                ('FONTSIZE', (0, 0), (-1, 0), 10),
                                ('FONTSIZE', (0, 1), (-1, -1), 10),
                                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                                ('GRID', (0, 0), (-1, -1), 1, colors.black)
                            ]))
                            story.append(top5_churn_table)
                            story.append(Spacer(1, 0.3*inch))
                            
                            # Выводы и рекомендации
                            story.append(Paragraph("6. ВЫВОДЫ И РЕКОМЕНДАЦИИ", heading_style))
                            
                            avg_return = churn_table['Накопительный % возврата'].mean()
                            avg_churn = churn_table['Отток %'].mean()
                            
                            top5_size = sorted(diagonal_values.items(), key=lambda x: x[1], reverse=True)[:5]
                            conclusions = [
                                f"• Средний процент возврата клиентов: {avg_return:.1f}%",
                                f"• Средний процент оттока: {avg_churn:.1f}%",
                                f"• Наиболее стабильная когорта (по размеру): {top5_size[0][0]} ({int(top5_size[0][1])} клиентов)",
                                f"• Когорта с наилучшим возвратом: {churn_sorted_return.iloc[0, 0]} ({churn_sorted_return.iloc[0, 3]:.1f}%)",
                                f"• Когорта с наибольшим оттоком требует внимания: {churn_sorted_churn.iloc[0, 0]} ({churn_sorted_churn.iloc[0, 5]:.1f}%)"
                            ]
                            
                            for conclusion in conclusions:
                                story.append(Paragraph(conclusion, normal_style))
                                story.append(Spacer(1, 0.1*inch))
                            
                            # Собираем PDF
                            doc.build(story)
                            buffer.seek(0)
                            return buffer.getvalue()
                        
                        # Генерируем PDF при нажатии кнопки
                        pdf_data = create_analysis_pdf()
                        
                        with col_pdf_button:
                            st.download_button(
                                label="📊 Скачать анализ отчёта в PDF",
                                data=pdf_data,
                                file_name=f"анализ_когортный_{info['first_period']}_{info['last_period']}.pdf",
                                mime="application/pdf",
                                use_container_width=True,
                                key="download_analysis_pdf"
                            )
                else:
                    st.info("⏳ Загрузите файл и дождитесь завершения расчётов для генерации отчётов")
                
                # Отображение матрицы (только если данные готовы)
                if info:
                    # Уменьшаем отступ перед блоком матриц
                    st.markdown("<div style='margin-top: 5px;'></div>", unsafe_allow_html=True)
                    
                    # Добавляем CSS для компактного отображения таблицы без прокрутки
                    st.markdown("""
                    <style>
                    div[data-testid="stDataFrame"] > div {
                        overflow: visible !important;
                    }
                    div[data-testid="stDataFrame"] table {
                        font-size: 0.7rem !important;
                        width: 100% !important;
                    }
                    /* Убираем overflow с внутренних контейнеров таблицы */
                    div[data-testid="stDataFrame"] > div > div {
                        overflow: visible !important;
                    }
                    div[data-testid="stDataFrame"] th, 
                    div[data-testid="stDataFrame"] td {
                        padding: 0.2rem 0.4rem !important;
                        font-size: 0.7rem !important;
                        white-space: nowrap !important;
                        text-align: center !important;
                    }
                    div[data-testid="stDataFrame"] table th,
                    div[data-testid="stDataFrame"] table td {
                        text-align: center !important;
                    }
                    </style>
                    """, unsafe_allow_html=True)
                    
                    # Объединенный блок с переключателем отображения
                    # CSS стили для красивого оформления блока
                    st.markdown("""
                    <style>
                    /* Стили для блока с таблицей */
                    .matrix-block-container {
                        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                        padding: 20px;
                        border-radius: 15px;
                        box-shadow: 0 10px 30px rgba(0,0,0,0.2);
                        margin-bottom: 20px;
                    }
                    
                    /* Стили для кнопок переключения */
                    .stRadio > div {
                        background: transparent;
                        padding: 0;
                        border-radius: 0;
                        box-shadow: none;
                        border: none;
                        display: flex;
                        flex-direction: row;
                        gap: 10px;
                        align-items: stretch;
                    }
                    
                    .stRadio > div > label {
                        background: white !important;
                        color: #333 !important;
                        padding: 12px 8px !important;
                        border-radius: 8px !important;
                        margin: 0 !important;
                        font-weight: 700 !important;
                        font-size: 0.75rem !important;
                        line-height: 1.2 !important;
                        transition: all 0.3s ease !important;
                        border: 2px solid #ccc !important;
                        box-shadow: 0 2px 4px rgba(0,0,0,0.1) !important;
                        cursor: pointer !important;
                        text-align: center !important;
                        flex: 1 1 0 !important;
                        min-width: 0 !important;
                        min-height: 50px !important;
                        height: auto !important;
                        max-height: 60px !important;
                        display: flex !important;
                        align-items: center !important;
                        justify-content: center !important;
                        white-space: normal !important;
                        word-wrap: break-word !important;
                        overflow: hidden !important;
                    }
                    
                    .stRadio > div > label:hover {
                        transform: translateY(-2px) !important;
                        box-shadow: 0 4px 8px rgba(0,0,0,0.15) !important;
                        background: #f5f5f5 !important;
                        border-color: #999 !important;
                    }
                    
                    .stRadio > div > label[data-baseweb="radio"]:has(input:checked) {
                        background: white !important;
                        color: #333 !important;
                        border-color: #666 !important;
                        box-shadow: 0 4px 10px rgba(0, 0, 0, 0.2) !important;
                    }
                    
                    .stRadio input[type="radio"]:checked + label {
                        background: white !important;
                        color: #333 !important;
                    }
                    
                    /* Стили для таблицы - только базовое оформление, не мешаем встроенному тулбару */
                    div[data-testid="stDataFrame"] {
                        background: white;
                        border-radius: 10px;
                        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
                        max-width: 100% !important;
                    }
                    
                    /* Стили для блока кодов клиентов */
                    .clients-block {
                        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                        padding: 15px;
                        border-radius: 10px;
                        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
                    }
                    
                    /* Стили для описания */
                    .description-block {
                        background: transparent;
                        padding: 15px;
                        border-radius: 10px;
                        margin-bottom: 15px;
                        box-shadow: none;
                        color: inherit;
                    }
                    
                    /* Контейнер для кнопок - ограничиваем ширину как у описания */
                    .stRadio {
                        max-width: 100%;
                    }
                    
                    /* Ограничиваем ширину контейнера кнопок и растягиваем на всю ширину */
                    div[data-testid="stRadio"] {
                        max-width: 100% !important;
                        width: 100% !important;
                    }
                    
                    /* Растягиваем контейнер с кнопками на всю ширину колонки */
                    .stRadio > div {
                        width: 100% !important;
                        display: flex !important;
                        flex-direction: row !important;
                        gap: 10px !important;
                    }
                    
                    /* Стили для кнопок Excel и PDF - светлые белые с большим шрифтом */
                    div[data-testid="stDownloadButton"] button,
                    div[data-testid="stButton"] button {
                        background: #f8f9fa !important;
                        color: #333 !important;
                        padding: 15px 20px !important;
                        border-radius: 8px !important;
                        margin: 0 !important;
                        font-weight: 700 !important;
                        font-size: 1.1rem !important;
                        line-height: 1.3 !important;
                        transition: all 0.3s ease !important;
                        border: 2px solid #e0e0e0 !important;
                        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05) !important;
                        cursor: pointer !important;
                        text-align: center !important;
                        min-height: 60px !important;
                        height: auto !important;
                        display: flex !important;
                        align-items: center !important;
                        justify-content: center !important;
                        white-space: normal !important;
                        word-wrap: break-word !important;
                        width: 100% !important;
                    }
                    
                    div[data-testid="stDownloadButton"] button:hover,
                    div[data-testid="stButton"] button:hover {
                        transform: translateY(-2px) !important;
                        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1) !important;
                        background: #ffffff !important;
                        border-color: #d0d0d0 !important;
                    }
                    
                    div[data-testid="stDownloadButton"] button:active,
                    div[data-testid="stButton"] button:active {
                        transform: translateY(0) !important;
                        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05) !important;
                    }
                    </style>
                    """, unsafe_allow_html=True)
                    
                    # Создаем колонки для выравнивания кнопок с блоком описания
                    # Кнопки занимают всю ширину до блока кодов клиентов (соотношение 4:1 как у таблицы)
                    col_buttons_container, col_empty = st.columns([4, 1])
                    
                    with col_buttons_container:
                        # Переключатель для выбора типа отображения (горизонтально, на уровне с таблицей)
                        view_type = st.radio(
                            "",
                            options=[
                                "Динамика уникальных клиентов когорт",
                                "Динамика накопления возврата",
                                "Динамика накопления возврата в %",
                                "Приток возврата в %",
                                "Отток клиентов из категории"
                            ],
                            horizontal=True,
                            key="view_type_selector"
                        )
                    
                    # Уменьшаем отступ между кнопками и таблицей
                    st.markdown("<div style='margin-top: 5px;'></div>", unsafe_allow_html=True)
                    
                    # Основной контент
                    # Инициализируем переменные для таблицы и описания
                    display_matrix = None
                    description_text = ""
                    view_key = ""
                    
                    # Подготовка данных в зависимости от выбранного типа
                    if view_type == "Динамика уникальных клиентов когорт":
                        # Применяем цветовое форматирование
                        matrix_int = cohort_matrix.astype(int)
                        display_matrix = apply_matrix_color_gradient(matrix_int.astype(float), horizontal_dynamics=True, hide_before_diagonal=True)
                        display_matrix = display_matrix.format(precision=0, thousands=',', decimal='.')
                        description_text = "Диагональ показывает количество уникальных клиентов в каждом периоде. Пересечения показывают количество клиентов, которые были активны в обоих периодах."
                        view_key = "cohort"
                        
                    elif view_type == "Динамика накопления возврата":
                        accumulation_matrix = st.session_state.accumulation_matrix
                        matrix_int_accum = accumulation_matrix.astype(int)
                        display_matrix = apply_matrix_color_gradient(matrix_int_accum.astype(float), hide_zeros=True)
                        display_matrix = display_matrix.format(precision=0, thousands=',', decimal='.')
                        description_text = "Показывает накопление уникальных клиентов когорты по периодам. Каждая ячейка содержит количество уникальных клиентов когорты, которые вернулись в любой период от начала когорты до текущего включительно."
                        view_key = "accumulation"
                        
                    elif view_type == "Динамика накопления возврата в %":
                        accumulation_percent_matrix = st.session_state.accumulation_percent_matrix
                        display_matrix = apply_matrix_color_gradient(accumulation_percent_matrix, hide_zeros=True, horizontal_dynamics=True, hide_before_diagonal=True)
                        
                        # Форматирование процентов
                        def format_percent_cell(val):
                            if pd.isna(val) or val == '':
                                return ''
                            try:
                                val_float = float(val)
                                if val_float == 0:
                                    return ''
                                return f"{val_float:.1f}%"
                            except (ValueError, TypeError):
                                if isinstance(val, str) and '%' in val:
                                    return val
                                return ''
                        
                        display_matrix = display_matrix.format(formatter=format_percent_cell)
                        description_text = "Показывает долю накопления уникальных клиентов когорты от общего количества клиентов в когорте. Значения выражены в процентах."
                        view_key = "accumulation_percent"
                        
                    elif view_type == "Приток возврата в %":
                        inflow_matrix = st.session_state.inflow_matrix
                        display_matrix = apply_matrix_color_gradient(inflow_matrix, hide_zeros=True, horizontal_dynamics=True, hide_before_diagonal=True)
                        
                        # Форматирование процентов для притока
                        def format_inflow_percent_cell(val):
                            if pd.isna(val) or val == '':
                                return ''
                            try:
                                val_float = float(val)
                                if val_float == 0:
                                    return ''
                                return f"{val_float:.1f}%"
                            except (ValueError, TypeError):
                                if isinstance(val, str) and '%' in val:
                                    return val
                                return ''
                        
                        # Добавляем 0.0% на диагонали
                        for row_name in display_matrix.data.index:
                            if row_name in display_matrix.data.columns:
                                display_matrix.data.loc[row_name, row_name] = '0.0%'
                        
                        format_dict_inflow = {col: format_inflow_percent_cell for col in display_matrix.data.columns}
                        display_matrix = display_matrix.format(format_dict_inflow)
                        description_text = "Показывает прирост уникальных клиентов когорты между периодами. Диагональ = 0%, первый период после диагонали = процент возврата, остальные = разница между накопительными процентами соседних периодов."
                        view_key = "inflow"
                    
                    elif view_type == "Отток клиентов из категории":
                        # Используем сохраненную таблицу оттока
                        if st.session_state.get('churn_table') is not None:
                            churn_table = st.session_state.churn_table
                            
                            # Форматируем таблицу для отображения
                            churn_display = churn_table.copy()
                            churn_display['Накопительный % возврата'] = churn_display['Накопительный % возврата'].apply(lambda x: f"{x:.1f}%")
                            churn_display['Отток %'] = churn_display['Отток %'].apply(lambda x: f"{x:.1f}%")
                            
                            # Убеждаемся, что когорта - первый столбец
                            column_order = ['Когорта', 'Кол-во клиентов когорты', 'Накопительное кол-во возврата', 
                                          'Накопительный % возврата', 'Отток кол-во', 'Отток %']
                            churn_display = churn_display[column_order]
                            
                            # Применяем стили для центрирования значений во всех столбцах
                            def center_format(val):
                                return 'text-align: center'
                            
                            # Создаем стилизованную таблицу с центрированием
                            styled_churn = churn_display.style.applymap(center_format)
                            
                            # Используем styled_churn как display_matrix для единообразия
                            display_matrix = styled_churn
                            description_text = "Показывает клиентов, которые не вернулись в категорию ни разу после периода когорты."
                            view_key = "churn"
                        else:
                            st.error("Таблица оттока не загружена. Пожалуйста, загрузите данные заново.")
                            display_matrix = None
                            description_text = ""
                            view_key = ""
                    
                    # Отображение описания с красивым оформлением
                    if description_text:
                        st.markdown(f'<div class="description-block">{description_text}</div>', unsafe_allow_html=True)
                    
                    # Создаем колонки для таблицы и кодов клиентов
                    col_table, col_clients = st.columns([4, 1])
                    
                    with col_table:
                        # Отображение таблицы (широкая) с поддержкой полноэкранного режима
                        if display_matrix is not None:
                            # Для таблицы оттока скрываем индекс
                            if view_key == "churn":
                                st.dataframe(
                                    display_matrix,
                                    use_container_width=True,
                                    hide_index=True
                                )
                                # Добавляем CSS для центрирования значений в таблице оттока
                                st.markdown("""
                                <style>
                                div[data-testid="stDataFrame"] table td {
                                    text-align: center !important;
                                }
                                div[data-testid="stDataFrame"] table th {
                                    text-align: center !important;
                                }
                                </style>
                                """, unsafe_allow_html=True)
                            else:
                                st.dataframe(
                                    display_matrix,
                                    use_container_width=True
                                )
                        else:
                            st.info("Выберите тип отображения для просмотра данных.")
                    
                    with col_clients:
                        # Компактный блок кодов клиентов
                        st.markdown('<div style="background: white; padding: 10px; border-radius: 8px; margin-bottom: 10px; border: 2px solid #ccc; box-shadow: 0 2px 4px rgba(0,0,0,0.1);"><h4 style="color: #333; margin: 0;">👥 Коды клиентов</h4></div>', unsafe_allow_html=True)
                        
                        # Коды клиентов в зависимости от выбранного типа
                        if view_key == "cohort":
                            selected_cohort = st.selectbox(
                                "Когорта:",
                                options=sorted_periods,
                                index=0,
                                help="Выберите период, когда клиенты впервые появились",
                                key="cohort_select_unified_1"
                            )
                            
                            selected_period = st.selectbox(
                                "Период:",
                                options=sorted_periods,
                                index=min(1, len(sorted_periods) - 1) if len(sorted_periods) > 1 else 0,
                                help="Выберите период, для которого нужно показать клиентов",
                                key="period_select_unified_1"
                            )
                            
                            if selected_cohort and selected_period:
                                period_clients_cache = st.session_state.get('period_clients_cache', None)
                                common_clients = get_cohort_clients(df, year_month_col, client_col, selected_cohort, selected_period, period_clients_cache)
                                
                                if common_clients:
                                    st.write(f"**Найдено: {len(common_clients)}**")
                                    clients_csv = "\n".join([str(client) for client in common_clients])
                                    create_copy_button(
                                        clients_csv,
                                        f"📋 Копировать ({len(common_clients)})",
                                        "copy_clients_unified_1"
                                    )
                                else:
                                    st.info(f"❌ Нет данных")
                        
                        elif view_key == "accumulation":
                            selected_cohort = st.selectbox(
                                "Когорта:",
                                options=sorted_periods,
                                index=0,
                                help="Выберите период когорты",
                                key="cohort_select_unified_2"
                            )
                            
                            selected_period = st.selectbox(
                                "Период:",
                                options=sorted_periods,
                                index=min(1, len(sorted_periods) - 1) if len(sorted_periods) > 1 else 0,
                                help="Выберите период, до которого показывать накопленных клиентов",
                                key="period_select_unified_2"
                            )
                            
                            if selected_cohort and selected_period:
                                period_clients_cache = st.session_state.get('period_clients_cache', None)
                                accumulation_clients = get_accumulation_clients(df, year_month_col, client_col, sorted_periods, selected_cohort, selected_period, period_clients_cache)
                                
                                if accumulation_clients:
                                    st.write(f"**Найдено: {len(accumulation_clients)}**")
                                    clients_csv = "\n".join([str(client) for client in accumulation_clients])
                                    create_copy_button(
                                        clients_csv,
                                        f"📋 Копировать ({len(accumulation_clients)})",
                                        "copy_clients_unified_2"
                                    )
                                else:
                                    st.info(f"❌ Нет данных")
                        
                        elif view_key == "accumulation_percent":
                            selected_cohort = st.selectbox(
                                "Когорта:",
                                options=sorted_periods,
                                index=0,
                                help="Выберите период когорты",
                                key="cohort_select_unified_3"
                            )
                            
                            selected_period = st.selectbox(
                                "Период:",
                                options=sorted_periods,
                                index=min(1, len(sorted_periods) - 1) if len(sorted_periods) > 1 else 0,
                                help="Выберите период, до которого показывать накопленных клиентов",
                                key="period_select_unified_3"
                            )
                            
                            if selected_cohort and selected_period:
                                period_clients_cache = st.session_state.get('period_clients_cache', None)
                                accumulation_clients = get_accumulation_clients(df, year_month_col, client_col, sorted_periods, selected_cohort, selected_period, period_clients_cache)
                                
                                if accumulation_clients:
                                    st.write(f"**Найдено: {len(accumulation_clients)}**")
                                    clients_csv = "\n".join([str(client) for client in accumulation_clients])
                                    create_copy_button(
                                        clients_csv,
                                        f"📋 Копировать ({len(accumulation_clients)})",
                                        "copy_clients_unified_3"
                                    )
                                else:
                                    st.info(f"❌ Нет данных")
                        
                        elif view_key == "inflow":
                            selected_cohort = st.selectbox(
                                "Когорта:",
                                options=sorted_periods,
                                index=0,
                                help="Выберите период когорты",
                                key="cohort_select_unified_4"
                            )
                            
                            selected_period = st.selectbox(
                                "Период:",
                                options=sorted_periods,
                                index=min(1, len(sorted_periods) - 1) if len(sorted_periods) > 1 else 0,
                                help="Выберите период, для которого показать новых вернувшихся клиентов",
                                key="period_select_unified_4"
                            )
                            
                            if selected_cohort and selected_period:
                                period_clients_cache = st.session_state.get('period_clients_cache', None)
                                inflow_clients = get_inflow_clients(df, year_month_col, client_col, sorted_periods, selected_cohort, selected_period, period_clients_cache)
                                
                                if inflow_clients:
                                    st.write(f"**Найдено: {len(inflow_clients)}**")
                                    clients_csv = "\n".join([str(client) for client in inflow_clients])
                                    create_copy_button(
                                        clients_csv,
                                        f"📋 Копировать ({len(inflow_clients)})",
                                        "copy_clients_unified_4"
                                    )
                                else:
                                    st.info(f"❌ Нет данных")
                        
                        elif view_key == "churn":
                            # Для оттока только выбор когорты, без периода
                            selected_cohort = st.selectbox(
                                "Когорта:",
                                options=sorted_periods,
                                index=0,
                                help="Выберите когорту для скачивания списка клиентов оттока из категории",
                                key="cohort_select_unified_5"
                            )
                            
                            if selected_cohort:
                                period_clients_cache = st.session_state.get('period_clients_cache', None)
                                client_cohorts_cache = st.session_state.get('client_cohorts_cache', None)
                                churn_clients = get_churn_clients(df, year_month_col, client_col, sorted_periods, selected_cohort, period_clients_cache, client_cohorts_cache)
                                
                                if churn_clients:
                                    st.write(f"**Найдено: {len(churn_clients)}**")
                                    clients_csv = "\n".join([str(client) for client in churn_clients])
                                    create_copy_button(
                                        clients_csv,
                                        f"📋 Копировать ({len(churn_clients)})",
                                        "copy_clients_unified_5"
                                    )
                                else:
                                    st.info(f"❌ Нет данных")
                                
                                # Кнопка для скачивания всех когорт (всегда видна)
                                all_churn_clients = set()
                                client_cohorts_cache = st.session_state.get('client_cohorts_cache', None)
                                for cohort in sorted_periods:
                                    cohort_churn = get_churn_clients(df, year_month_col, client_col, sorted_periods, cohort, period_clients_cache, client_cohorts_cache)
                                    all_churn_clients.update(cohort_churn)
                                
                                if all_churn_clients:
                                    all_clients_csv = "\n".join([str(client) for client in sorted(all_churn_clients)])
                                    create_copy_button(
                                        all_clients_csv,
                                        f"📋 Копировать коды клиентов оттока всех когорт ({len(all_churn_clients)})",
                                        "copy_all_churn_clients"
                                    )
                    
                    # Шестой блок - Присутствие клиентов оттока в других категориях
                    st.markdown("---")
                    
                    # Блок с заголовками, инструкциями и изображением в одной строке
                    col_churn_title_left, col_churn_title_right = st.columns([1, 1])
                    
                    with col_churn_title_left:
                        st.subheader("🔍 Присутствие клиентов оттока в других категориях и отток из сети")
                        # Текст инструкций прямо под заголовком
                        st.markdown("""
                        1) Скачайте коды клиентов оттока всех когорт в блоке "Отток клиентов из категории"
                        
                        2) Примените коды клиентов. Отберите анализируемый период и необходимые категории.
                        
                        3) Зайдите на лист "Конструктор" и выведите отчёт по шаблону справа.
                        
                        4) Скачайте документ в Qlik и загрузите в ячейку справа.
                        """)
                    
                    with col_churn_title_right:
                        st.subheader("📋 Шаблон загрузки данных из Qlik")
                        # Пытаемся найти скриншот шаблона для категорий (изображение сразу под заголовком)
                        churn_categories_image_paths = [
                            'qlik_template_categories.png',
                            'qlik_template_categories.jpg',
                            'qlik_template_categories.jpeg',
                            'шаблон_qlik_категории.png',
                            'шаблон_qlik_категории.jpg',
                            'шаблон_qlik_категории.jpeg',
                            'churn_categories_template.png',
                            'churn_categories_template.jpg',
                            'churn_categories_template.jpeg'
                        ]
                        image_found = False
                        for img_path in churn_categories_image_paths:
                            if os.path.exists(img_path):
                                st.image(img_path, use_container_width=True)
                                image_found = True
                                break
                        if not image_found:
                            st.info("📸 Поместите скриншот шаблона загрузки данных из Qlik в папку проекта с одним из имён: qlik_template_categories.png, шаблон_qlik_категории.png или churn_categories_template.png")
                        
                        # Загрузчик Excel файла прямо под изображением
                        uploaded_file_categories = st.file_uploader(
                            "Выберите Excel файл с данными о присутствии клиентов оттока в других категориях",
                            type=['xlsx', 'xls'],
                            help="Загрузите файл, скачанный из Qlik согласно шаблону выше",
                            key="upload_categories_file"
                        )
                    
                    # Обработка загруженного файла
                    if uploaded_file_categories is not None:
                        try:
                            # Загрузка Excel файла
                            if uploaded_file_categories.name.endswith('.xlsx'):
                                df_categories = pd.read_excel(uploaded_file_categories, engine='openpyxl')
                            else:
                                df_categories = pd.read_excel(uploaded_file_categories, engine='xlrd')
                            
                            # Определяем столбцы
                            group_col = None
                            year_month_col = None
                            month_col = None
                            clients_col = None
                            client_code_col = None
                            
                            # Ищем столбец Группа (может быть Группа1, Группа2, Группа3 и т.д.)
                            for col in df_categories.columns:
                                col_lower = str(col).lower().strip()
                                if 'группа' in col_lower:
                                    group_col = col
                                    break
                            
                            # Ищем столбец Год-месяц
                            for col in df_categories.columns:
                                col_lower = str(col).lower().strip()
                                if ('год' in col_lower and 'месяц' in col_lower) or ('год-месяц' in col_lower):
                                    year_month_col = col
                                    break
                            
                            # Ищем столбец месяц
                            for col in df_categories.columns:
                                col_lower = str(col).lower().strip()
                                if col_lower == 'месяц' or (col_lower.startswith('месяц') and len(col_lower.split()) == 1):
                                    month_col = col
                                    break
                            
                            # Ищем столбец Клиентов
                            for col in df_categories.columns:
                                col_lower = str(col).lower().strip()
                                if 'клиент' in col_lower and ('ов' in col_lower or 'ов' in col_lower):
                                    clients_col = col
                                    break
                            
                            # Ищем столбец Код клиента
                            for col in df_categories.columns:
                                col_lower = str(col).lower().strip()
                                if 'код' in col_lower and 'клиент' in col_lower:
                                    client_code_col = col
                                    break
                            
                            # Проверяем наличие всех необходимых столбцов
                            if group_col is None:
                                st.error("❌ Не найден столбец с категориями (Группа1, Группа2, Группа3 и т.д.). Убедитесь, что в файле есть столбец с названием, содержащим 'Группа'.")
                            elif client_code_col is None:
                                st.error("❌ Не найден столбец 'Код клиента'. Убедитесь, что в файле есть столбец с названием, содержащим 'Код' и 'клиент'.")
                            elif year_month_col is None:
                                st.warning("⚠️ Не найден столбец 'Год-месяц'. Данные будут обработаны без фильтрации по периоду.")
                            else:
                                # Получаем уникальные категории
                                categories = df_categories[group_col].dropna().unique()
                                categories = sorted([str(cat) for cat in categories if str(cat).strip() != ''])
                                
                                # Сохраняем данные о категориях в session_state для использования в Excel отчёте
                                st.session_state.df_categories = df_categories
                                st.session_state.categories_list = categories
                                st.session_state.group_col_name = group_col
                                st.session_state.year_month_col_name = year_month_col
                                st.session_state.client_code_col_name = client_code_col
                                
                                # Получаем клиентов оттока для каждой когорты
                                period_clients_cache = st.session_state.get('period_clients_cache', None)
                                
                                # Рассчитываем метрики для всех когорт для сводной таблицы
                                total_present_after_cohort_by_cohort = {}
                                total_present_after_cohort_percent_by_cohort = {}
                                network_churn_by_cohort = {}
                                network_churn_percent_by_cohort = {}
                                
                                # Собираем всех клиентов из категорий (для всех периодов)
                                all_category_clients_all_periods = set()
                                if year_month_col is not None:
                                    for category in categories:
                                        category_data = df_categories[df_categories[group_col] == category]
                                        category_clients = set(category_data[client_code_col].dropna().astype(str).unique())
                                        all_category_clients_all_periods.update(category_clients)
                                else:
                                    for category in categories:
                                        category_data = df_categories[df_categories[group_col] == category]
                                        category_clients = set(category_data[client_code_col].dropna().astype(str).unique())
                                        all_category_clients_all_periods.update(category_clients)
                                
                                # Для каждой когорты рассчитываем метрики
                                churn_table = st.session_state.churn_table
                                client_cohorts_cache = st.session_state.get('client_cohorts_cache', None)
                                # Собираем всех клиентов оттока из сети для всех когорт
                                all_network_churn_clients = set()
                                for cohort_period in sorted_periods:
                                    # Получаем клиентов оттока для этой когорты
                                    churn_clients_set_cohort = set(get_churn_clients(df, year_month_col, client_col, sorted_periods, cohort_period, period_clients_cache, client_cohorts_cache))
                                    churn_clients_set_cohort = {str(client) for client in churn_clients_set_cohort}
                                    
                                    # Получаем отток из категории для этой когорты
                                    cohort_row = churn_table[churn_table['Когорта'] == cohort_period]
                                    churn_count_cohort = int(cohort_row.iloc[0]['Отток кол-во']) if not cohort_row.empty else 0
                                    cohort_size_cohort = int(cohort_row.iloc[0]['Кол-во клиентов когорты']) if not cohort_row.empty else 0
                                    
                                    # Определяем периоды начиная с этой когорты
                                    cohort_index_cohort = sorted_periods.index(cohort_period) if cohort_period in sorted_periods else 0
                                    periods_from_cohort_cohort = sorted_periods[cohort_index_cohort:]
                                    # Периоды ПОСЛЕ когорты (исключая период когорты)
                                    periods_after_cohort_cohort = periods_from_cohort_cohort[1:] if len(periods_from_cohort_cohort) > 1 else []
                                    
                                    # Клиенты оттока, присутствующие в других категориях ПОСЛЕ месяца когорты
                                    all_category_clients_after_cohort = set()
                                    if year_month_col is not None and len(periods_after_cohort_cohort) > 0:
                                        for category in categories:
                                            category_data = df_categories[df_categories[group_col] == category]
                                            category_data_filtered = category_data[category_data[year_month_col].isin(periods_after_cohort_cohort)]
                                            category_clients = set(category_data_filtered[client_code_col].dropna().astype(str).unique())
                                            all_category_clients_after_cohort.update(category_clients)
                                    elif year_month_col is None:
                                        all_category_clients_after_cohort = all_category_clients_all_periods
                                    
                                    present_in_categories_after_cohort = churn_clients_set_cohort & all_category_clients_after_cohort
                                    total_present_after_cohort_by_cohort[cohort_period] = len(present_in_categories_after_cohort)
                                    
                                    # % присутствия после месяца когорты
                                    present_after_cohort_percent = (len(present_in_categories_after_cohort) / cohort_size_cohort * 100) if cohort_size_cohort > 0 else 0
                                    total_present_after_cohort_percent_by_cohort[cohort_period] = present_after_cohort_percent
                                    
                                    # Отток из сети = Отток из категории - Клиентов когорты присутствуют в других категориях после месяца когорты
                                    network_churn_cohort = churn_count_cohort - len(present_in_categories_after_cohort)
                                    network_churn_by_cohort[cohort_period] = max(0, network_churn_cohort)  # Не может быть отрицательным
                                    
                                    # % оттока из сети
                                    network_churn_percent_cohort = (network_churn_by_cohort[cohort_period] / cohort_size_cohort * 100) if cohort_size_cohort > 0 else 0
                                    network_churn_percent_by_cohort[cohort_period] = network_churn_percent_cohort
                                    
                                    # Собираем клиентов оттока из сети для этой когорты
                                    network_churn_clients_cohort = churn_clients_set_cohort - all_category_clients_after_cohort
                                    all_network_churn_clients.update(network_churn_clients_cohort)
                                
                                # Создаем таблицу для сохранения в session_state
                                summary_table_excel = pd.DataFrame({
                                    'Отток из сети': network_churn_by_cohort,
                                    'Доля оттока из сети от когорты': network_churn_percent_by_cohort,
                                    'Итого присутствуют в других категориях после месяца когорты': total_present_after_cohort_by_cohort,
                                    'Доля присутствуют в других категориях после месяца когорты': total_present_after_cohort_percent_by_cohort
                                })
                                summary_table_excel = summary_table_excel.T
                                
                                # Сохраняем данные для Excel отчёта и сводной таблицы
                                st.session_state.category_summary_table = summary_table_excel
                                st.session_state.category_cohort_table = None
                                
                                # Обновляем Excel отчёт после сохранения всех данных
                                if 'excel_report_cache_key' in st.session_state:
                                    del st.session_state.excel_report_cache_key
                                
                                # Перегенерируем Excel отчёт после сохранения данных о категориях
                                # Используем st.rerun() для обновления, но только если данные изменились
                                # Вместо этого просто перегенерируем отчет
                                try:
                                    # Небольшая задержка для гарантии сохранения данных
                                    st.session_state.excel_report_data = create_full_report_excel()
                                except Exception as e:
                                    st.warning(f"Не удалось обновить Excel отчёт: {str(e)}")
                                
                                # Новый интерфейс: слева выбор когорты, справа таблица
                                st.markdown("### 📊 Присутствие клиентов оттока когорты в других категориях товаров")
                                
                                col_cohort_select, col_table = st.columns([1, 4])
                                
                                with col_cohort_select:
                                    selected_cohort = st.selectbox(
                                        "Выберите когорту:",
                                        options=sorted_periods,
                                        index=0,
                                        help="Выберите когорту для анализа присутствия её клиентов оттока в других категориях",
                                        key="category_cohort_select"
                                    )
                                    
                                    # Определяем периоды начиная с выбранной когорты
                                    cohort_index = sorted_periods.index(selected_cohort) if selected_cohort in sorted_periods else 0
                                    periods_from_cohort = sorted_periods[cohort_index:]
                                    # Периоды ПОСЛЕ когорты (исключая период когорты) - начинаем расчет с этого периода
                                    periods_after_cohort = periods_from_cohort[1:] if len(periods_from_cohort) > 1 else []
                                    
                                    # Получаем клиентов оттока для выбранной когорты
                                    client_cohorts_cache = st.session_state.get('client_cohorts_cache', None)
                                    churn_clients_set = set(get_churn_clients(df, year_month_col, client_col, sorted_periods, selected_cohort, period_clients_cache, client_cohorts_cache))
                                    churn_clients_set = {str(client) for client in churn_clients_set}
                                    
                                    # Получаем размер когорты и отток из churn_table
                                    churn_table = st.session_state.churn_table
                                    cohort_row = churn_table[churn_table['Когорта'] == selected_cohort]
                                    cohort_size = int(cohort_row.iloc[0]['Кол-во клиентов когорты']) if not cohort_row.empty else 0
                                    churn_count = int(cohort_row.iloc[0]['Отток кол-во']) if not cohort_row.empty else 0
                                    
                                    # Клиенты оттока, присутствующие в других категориях ПОСЛЕ месяца когорты
                                    all_category_clients_after_cohort = set()
                                    if year_month_col is not None and len(periods_after_cohort) > 0:
                                        # Используем только данные из периодов ПОСЛЕ выбранной когорты
                                        for category in categories:
                                            category_data = df_categories[df_categories[group_col] == category]
                                            # Фильтруем только периоды ПОСЛЕ выбранной когорты
                                            category_data_filtered = category_data[category_data[year_month_col].isin(periods_after_cohort)]
                                            category_clients = set(category_data_filtered[client_code_col].dropna().astype(str).unique())
                                            all_category_clients_after_cohort.update(category_clients)
                                    elif year_month_col is None:
                                        # Если нет столбца "Год-месяц", собираем всех клиентов из всех категорий
                                        for category in categories:
                                            category_data = df_categories[df_categories[group_col] == category]
                                            category_clients = set(category_data[client_code_col].dropna().astype(str).unique())
                                            all_category_clients_after_cohort.update(category_clients)
                                    
                                    present_in_categories_after_cohort = churn_clients_set & all_category_clients_after_cohort
                                    present_count_after_cohort = len(present_in_categories_after_cohort)
                                    present_percent_after_cohort = (present_count_after_cohort / cohort_size * 100) if cohort_size > 0 else 0
                                    
                                    # Отток из сети = Отток из категории - Клиентов когорты присутствуют в других категориях после месяца когорты
                                    network_churn = churn_count - present_count_after_cohort
                                    network_churn = max(0, network_churn)  # Не может быть отрицательным
                                    network_churn_percent = (network_churn / cohort_size * 100) if cohort_size > 0 else 0
                                    
                                    # Вычисляем клиентов оттока из сети
                                    # Это клиенты оттока, которые НЕ присутствуют в других категориях после месяца когорты
                                    network_churn_clients = churn_clients_set - all_category_clients_after_cohort
                                    network_churn_clients_list = sorted(list(network_churn_clients))
                                    
                                    # Выводим метрики в текстовом формате с цветом (проценты в скобках)
                                    metrics_html = f"""
                                    <div style="line-height: 2;">
                                    <p style="color: #333; font-size: 1rem; margin: 8px 0;">
                                        <strong style="color: #1f77b4;">Клиентов когорты присутствуют в других категориях после месяца когорты:</strong> 
                                        <span style="color: #2c3e50; font-weight: 600;">{present_count_after_cohort} ({present_percent_after_cohort:.1f}%)</span>
                                    </p>
                                    <p style="color: #333; font-size: 1rem; margin: 8px 0;">
                                        <strong style="color: #1f77b4;">Отток из сети:</strong> 
                                        <span style="color: #e74c3c; font-weight: 600;">{network_churn} ({network_churn_percent:.1f}%)</span>
                                    </p>
                                    </div>
                                    """
                                    st.markdown(metrics_html, unsafe_allow_html=True)
                                    
                                    # Кнопка копирования кодов клиентов оттока из сети для выбранной когорты
                                    if network_churn_clients_list:
                                        network_churn_clients_csv = "\n".join([str(client) for client in network_churn_clients_list])
                                        create_copy_button(
                                            network_churn_clients_csv,
                                            f"📋 Копировать коды клиентов оттока из сети ({len(network_churn_clients_list)})",
                                            f"copy_network_churn_{selected_cohort}"
                                        )
                                    else:
                                        st.info("ℹ️ Отток из сети равен 0 или все клиенты оттока присутствуют в других категориях")
                                    
                                    # Кнопка копирования всех кодов клиентов оттока из сети (для всех когорт)
                                    if all_network_churn_clients:
                                        all_network_churn_clients_list = sorted(list(all_network_churn_clients))
                                        all_network_churn_clients_csv = "\n".join([str(client) for client in all_network_churn_clients_list])
                                        create_copy_button(
                                            all_network_churn_clients_csv,
                                            f"📋 Копировать все коды клиентов оттока из сети ({len(all_network_churn_clients_list)})",
                                            "copy_all_network_churn_clients"
                                        )
                                
                                with col_table:
                                    # Определяем периоды ПОСЛЕ выбранной когорты (для использования в таблице)
                                    cohort_index_table = sorted_periods.index(selected_cohort) if selected_cohort in sorted_periods else 0
                                    periods_from_cohort_table = sorted_periods[cohort_index_table:]
                                    # Периоды ПОСЛЕ когорты (исключая период когорты) - начинаем расчет с этого периода
                                    periods_after_cohort_table = periods_from_cohort_table[1:] if len(periods_from_cohort_table) > 1 else []
                                    
                                    # Создаем таблицу: категории по строкам, периоды по столбцам (только ПОСЛЕ выбранной когорты)
                                    category_period_table = pd.DataFrame(index=categories, columns=periods_after_cohort_table)
                                    
                                    # Словарь для хранения уникальных клиентов по периодам (для итоговой строки)
                                    period_unique_clients = {period: set() for period in periods_after_cohort_table}
                                    
                                    # Словарь для хранения уникальных клиентов по категориям (для итогового столбца)
                                    category_unique_clients = {category: set() for category in categories}
                                    
                                    # Если есть столбец "Год-месяц", используем его для фильтрации по периодам
                                    if year_month_col is not None:
                                        # Для каждого периода ПОСЛЕ выбранной когорты проверяем присутствие клиентов оттока в категориях
                                        for period in periods_after_cohort_table:
                                            # Фильтруем данные по периоду
                                            period_data = df_categories[df_categories[year_month_col] == period]
                                            
                                            # Для каждой категории считаем количество клиентов оттока, присутствующих в этом периоде
                                            for category in categories:
                                                # Данные категории в этом периоде
                                                category_period_data = period_data[
                                                    (period_data[group_col] == category) & 
                                                    (period_data[client_code_col].notna())
                                                ]
                                                
                                                # Коды клиентов этой категории в этом периоде
                                                category_period_clients = set(
                                                    category_period_data[client_code_col].dropna().astype(str).unique()
                                                )
                                                
                                                # Находим пересечение: клиенты оттока выбранной когорты, которые есть в этой категории в этом периоде
                                                intersection = churn_clients_set & category_period_clients
                                                category_period_table.loc[category, period] = len(intersection)
                                                
                                                # Добавляем в множества для итогов
                                                period_unique_clients[period].update(intersection)
                                                category_unique_clients[category].update(intersection)
                                    else:
                                        # Если нет столбца "Год-месяц", используем все данные без фильтрации по периоду
                                        # Создаем словарь: категория -> множество кодов клиентов
                                        category_clients_dict = {}
                                        for category in categories:
                                            category_data = df_categories[df_categories[group_col] == category]
                                            client_codes = set(category_data[client_code_col].dropna().astype(str).unique())
                                            category_clients_dict[category] = client_codes
                                        
                                        # Для каждого периода ПОСЛЕ выбранной когорты используем одинаковые данные
                                        for period in periods_after_cohort_table:
                                            for category in categories:
                                                category_clients_set = category_clients_dict.get(category, set())
                                                intersection = churn_clients_set & category_clients_set
                                                category_period_table.loc[category, period] = len(intersection)
                                                
                                                # Добавляем в множества для итогов
                                                period_unique_clients[period].update(intersection)
                                                category_unique_clients[category].update(intersection)
                                    
                                    # Заполняем NaN нулями
                                    category_period_table = category_period_table.fillna(0).astype(int)
                                    
                                    # Создаем итоговую строку по периодам (уникальные клиенты по всем категориям)
                                    totals_row = pd.Series(
                                        {period: len(period_unique_clients[period]) for period in periods_after_cohort_table},
                                        name='Итого клиентов'
                                    )
                                    
                                    # Создаем итоговый столбец по категориям (уникальные клиенты за весь период)
                                    totals_col = pd.Series(
                                        {category: len(category_unique_clients[category]) for category in categories},
                                        name='Итого'
                                    )
                                    
                                    # Добавляем итоговую строку в таблицу
                                    category_period_table_with_totals = category_period_table.copy()
                                    category_period_table_with_totals.loc['Итого клиентов'] = totals_row
                                    
                                    # Добавляем итоговый столбец
                                    category_period_table_with_totals['Итого'] = totals_col
                                    category_period_table_with_totals.loc['Итого клиентов', 'Итого'] = len(present_in_categories_after_cohort)
                                    
                                    # Переупорядочиваем строки: итоговая строка наверх
                                    new_index = ['Итого клиентов'] + [cat for cat in categories]
                                    category_period_table_with_totals = category_period_table_with_totals.reindex(new_index)
                                    
                                    # Переупорядочиваем столбцы: итоговый столбец слева (после индекса, перед периодами)
                                    new_columns = ['Итого'] + list(periods_after_cohort_table)
                                    category_period_table_with_totals = category_period_table_with_totals[new_columns]
                                    
                                    # Отображаем основную таблицу с итогами
                                    st.dataframe(
                                        category_period_table_with_totals,
                                        use_container_width=True
                                    )
                                    
                                    # Добавляем стили для центрирования, выделения итоговых значений жирным, пастельным цветом и закрепления
                                    st.markdown("""
                                    <style>
                                    div[data-testid="stDataFrame"] table td {
                                        text-align: center !important;
                                    }
                                    div[data-testid="stDataFrame"] table th {
                                        text-align: center !important;
                                    }
                                    /* Закрепляем первую строку (итоговая строка "Итого клиентов") сверху */
                                    div[data-testid="stDataFrame"] table tbody tr:first-child td,
                                    div[data-testid="stDataFrame"] table tbody tr:first-child th {
                                        font-weight: bold !important;
                                        background-color: #E3F2FD !important;
                                        position: sticky !important;
                                        top: 0 !important;
                                        z-index: 10 !important;
                                    }
                                    /* Закрепляем первый столбец данных (итоговый столбец "Итого") слева */
                                    div[data-testid="stDataFrame"] table tbody tr td:nth-child(2),
                                    div[data-testid="stDataFrame"] table thead tr th:nth-child(2) {
                                        font-weight: bold !important;
                                        background-color: #E3F2FD !important;
                                        position: sticky !important;
                                        left: 0 !important;
                                        z-index: 5 !important;
                                    }
                                    /* Закрепляем ячейку пересечения итоговых строки и столбца (и сверху, и слева) */
                                    div[data-testid="stDataFrame"] table tbody tr:first-child td:nth-child(2) {
                                        background-color: #BBDEFB !important;
                                        font-weight: bold !important;
                                        position: sticky !important;
                                        top: 0 !important;
                                        left: 0 !important;
                                        z-index: 15 !important;
                                    }
                                    /* Закрепляем заголовок итогового столбца */
                                    div[data-testid="stDataFrame"] table thead tr th:nth-child(2) {
                                        position: sticky !important;
                                        left: 0 !important;
                                        z-index: 6 !important;
                                    }
                                    </style>
                                    <script>
                                    // Дополнительный скрипт для гарантированного выделения жирным, цветом и закрепления
                                    setTimeout(function() {
                                        const tables = document.querySelectorAll('div[data-testid="stDataFrame"] table');
                                        tables.forEach(table => {
                                            // Первая строка (итоговая) - закрепляем сверху
                                            const firstRow = table.querySelector('tbody tr:first-child');
                                            if (firstRow) {
                                                firstRow.querySelectorAll('td, th').forEach(cell => {
                                                    cell.style.fontWeight = 'bold';
                                                    cell.style.position = 'sticky';
                                                    cell.style.top = '0';
                                                    cell.style.zIndex = '10';
                                                    if (!cell.style.backgroundColor || cell.style.backgroundColor === '') {
                                                        cell.style.backgroundColor = '#E3F2FD';
                                                    }
                                                });
                                            }
                                            // Первый столбец данных (итоговый) - закрепляем слева
                                            table.querySelectorAll('tbody tr').forEach(row => {
                                                const firstDataCell = row.querySelector('td:nth-child(2)');
                                                if (firstDataCell) {
                                                    firstDataCell.style.fontWeight = 'bold';
                                                    firstDataCell.style.position = 'sticky';
                                                    firstDataCell.style.left = '0';
                                                    firstDataCell.style.zIndex = '5';
                                                    if (!firstDataCell.style.backgroundColor || firstDataCell.style.backgroundColor === '') {
                                                        firstDataCell.style.backgroundColor = '#E3F2FD';
                                                    }
                                                }
                                            });
                                            const firstHeader = table.querySelector('thead th:nth-child(2)');
                                            if (firstHeader) {
                                                firstHeader.style.fontWeight = 'bold';
                                                firstHeader.style.backgroundColor = '#E3F2FD';
                                                firstHeader.style.position = 'sticky';
                                                firstHeader.style.left = '0';
                                                firstHeader.style.zIndex = '6';
                                            }
                                            // Ячейка пересечения - закрепляем и сверху, и слева
                                            const intersectionCell = table.querySelector('tbody tr:first-child td:nth-child(2)');
                                            if (intersectionCell) {
                                                intersectionCell.style.backgroundColor = '#BBDEFB';
                                                intersectionCell.style.position = 'sticky';
                                                intersectionCell.style.top = '0';
                                                intersectionCell.style.left = '0';
                                                intersectionCell.style.zIndex = '15';
                                            }
                                        });
                                    }, 100);
                                    </script>
                                    """, unsafe_allow_html=True)
                                
                        except Exception as e:
                            st.error(f"❌ Ошибка при обработке файла: {str(e)}")
                            st.exception(e)
                    
                    # Сводная таблица по всем когортам (после блока присутствия клиентов)
                    st.markdown("---")
                    st.subheader("📊 Сводная таблица по всем когортам")
                    
                    if st.session_state.get('churn_table') is not None:
                        churn_table = st.session_state.churn_table
                        
                        # Создаем сводную таблицу
                        summary_data = {}
                        
                        # 1. Кол-во клиентов в когорте
                        summary_data['Кол-во клиентов в когорте'] = {}
                        for _, row in churn_table.iterrows():
                            cohort = row['Когорта']
                            summary_data['Кол-во клиентов в когорте'][cohort] = int(row['Кол-во клиентов когорты'])
                        
                        # 2. Накопительное кол-во вернувшихся в категорию
                        summary_data['Накопительное кол-во вернувшихся в категорию'] = {}
                        for _, row in churn_table.iterrows():
                            cohort = row['Когорта']
                            summary_data['Накопительное кол-во вернувшихся в категорию'][cohort] = int(row['Накопительное кол-во возврата'])
                        
                        # 3. Накопительное кол-во вернувшихся в категорию %
                        summary_data['Накопительное кол-во вернувшихся в категорию %'] = {}
                        for _, row in churn_table.iterrows():
                            cohort = row['Когорта']
                            summary_data['Накопительное кол-во вернувшихся в категорию %'][cohort] = f"{row['Накопительный % возврата']:.1f}%"
                        
                        # 4. Отток из категории когорты
                        summary_data['Отток из категории когорты'] = {}
                        for _, row in churn_table.iterrows():
                            cohort = row['Когорта']
                            summary_data['Отток из категории когорты'][cohort] = int(row['Отток кол-во'])
                        
                        # 5. Отток из категории когорты %
                        summary_data['Отток из категории когорты %'] = {}
                        for _, row in churn_table.iterrows():
                            cohort = row['Когорта']
                            summary_data['Отток из категории когорты %'][cohort] = f"{row['Отток %']:.1f}%"
                        
                        # Инициализируем словари для метрик 6-9 заранее (заполняем нулями по умолчанию)
                        summary_data['Кол-во клиентов когорты в других категориях после месяца когорты'] = {}
                        summary_data['Кол-во клиентов когорты в других категориях после месяца когорты %'] = {}
                        summary_data['Отток из сети'] = {}
                        summary_data['Отток из сети %'] = {}
                        
                        for cohort in sorted_periods:
                            summary_data['Кол-во клиентов когорты в других категориях после месяца когорты'][cohort] = 0
                            summary_data['Кол-во клиентов когорты в других категориях после месяца когорты %'][cohort] = "0.0%"
                            summary_data['Отток из сети'][cohort] = 0
                            summary_data['Отток из сети %'][cohort] = "0.0%"
                        
                        # 6-9. Данные о присутствии в других категориях после месяца когорты и оттоке из сети (если есть данные)
                        if 'category_summary_table' in st.session_state and st.session_state.category_summary_table is not None:
                            category_summary = st.session_state.category_summary_table
                            
                            # 6. Кол-во клиентов когорты в других категориях после месяца когорты (обновляем значения)
                            if 'Итого присутствуют в других категориях после месяца когорты' in category_summary.index:
                                for cohort in sorted_periods:
                                    if cohort in category_summary.columns:
                                        value = category_summary.loc['Итого присутствуют в других категориях после месяца когорты', cohort]
                                        summary_data['Кол-во клиентов когорты в других категориях после месяца когорты'][cohort] = int(value) if pd.notna(value) else 0
                                    else:
                                        summary_data['Кол-во клиентов когорты в других категориях после месяца когорты'][cohort] = 0
                            
                            # 7. Кол-во клиентов когорты в других категориях после месяца когорты % (обновляем значения)
                            if 'Доля присутствуют в других категориях после месяца когорты' in category_summary.index:
                                for cohort in sorted_periods:
                                    if cohort in category_summary.columns:
                                        value = category_summary.loc['Доля присутствуют в других категориях после месяца когорты', cohort]
                                        if pd.notna(value):
                                            summary_data['Кол-во клиентов когорты в других категориях после месяца когорты %'][cohort] = f"{value:.1f}%"
                                        else:
                                            summary_data['Кол-во клиентов когорты в других категориях после месяца когорты %'][cohort] = "0.0%"
                                    else:
                                        summary_data['Кол-во клиентов когорты в других категориях после месяца когорты %'][cohort] = "0.0%"
                            else:
                                # Вычисляем процент вручную
                                for cohort in sorted_periods:
                                    cohort_size = summary_data['Кол-во клиентов в когорте'].get(cohort, 0)
                                    present_after_count = summary_data['Кол-во клиентов когорты в других категориях после месяца когорты'].get(cohort, 0)
                                    if cohort_size > 0:
                                        percent = (present_after_count / cohort_size) * 100
                                        summary_data['Кол-во клиентов когорты в других категориях после месяца когорты %'][cohort] = f"{percent:.1f}%"
                                    else:
                                        summary_data['Кол-во клиентов когорты в других категориях после месяца когорты %'][cohort] = "0.0%"
                            
                            # 8. Отток из сети (обновляем значения)
                            if 'Отток из сети' in category_summary.index:
                                for cohort in sorted_periods:
                                    if cohort in category_summary.columns:
                                        value = category_summary.loc['Отток из сети', cohort]
                                        summary_data['Отток из сети'][cohort] = int(value) if pd.notna(value) else 0
                                    else:
                                        summary_data['Отток из сети'][cohort] = 0
                            else:
                                for cohort in sorted_periods:
                                    summary_data['Отток из сети'][cohort] = 0
                            
                            # 9. Отток из сети % (обновляем значения)
                            if 'Доля оттока из сети от когорты' in category_summary.index:
                                for cohort in sorted_periods:
                                    if cohort in category_summary.columns:
                                        value = category_summary.loc['Доля оттока из сети от когорты', cohort]
                                        if pd.notna(value):
                                            summary_data['Отток из сети %'][cohort] = f"{value:.1f}%"
                                        else:
                                            summary_data['Отток из сети %'][cohort] = "0.0%"
                                    else:
                                        summary_data['Отток из сети %'][cohort] = "0.0%"
                            else:
                                # Если нет строки "Доля оттока из сети от когорты", вычисляем процент вручную
                                for cohort in sorted_periods:
                                    cohort_size = summary_data['Кол-во клиентов в когорте'].get(cohort, 0)
                                    network_churn = summary_data['Отток из сети'].get(cohort, 0)
                                    if cohort_size > 0:
                                        percent = (network_churn / cohort_size) * 100
                                        summary_data['Отток из сети %'][cohort] = f"{percent:.1f}%"
                                    else:
                                        summary_data['Отток из сети %'][cohort] = "0.0%"
                        # Если нет данных о категориях, значения уже заполнены нулями при инициализации выше
                        
                        # Создаем DataFrame
                        summary_df = pd.DataFrame(summary_data, index=sorted_periods).T
                        
                        # Отображаем таблицу
                        st.dataframe(
                            summary_df,
                            use_container_width=True
                        )
                        
                        # Добавляем стили для центрирования
                        st.markdown("""
                        <style>
                        div[data-testid="stDataFrame"] table td {
                            text-align: center !important;
                        }
                        div[data-testid="stDataFrame"] table th {
                            text-align: center !important;
                        }
                        </style>
                        """, unsafe_allow_html=True)
                    else:
                        st.info("Загрузите данные для отображения сводной таблицы")
                    
            except Exception as e:
                st.error(f"❌ Ошибка при построении матрицы: {str(e)}")
                st.exception(e)
        else:
            st.warning("⚠️ Необходимо указать столбцы для построения матрицы")
            
    except Exception as e:
        st.error(f"❌ Ошибка при загрузке файла: {str(e)}")
        st.session_state.uploaded_data = None
        st.session_state.df = None

