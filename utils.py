"""
Утилиты и вспомогательные функции
"""
import re
import pandas as pd
import streamlit.components.v1 as components
import json
from config import MONTHS_DICT


def parse_period(period_str):
    """Преобразует период в кортеж для сортировки.
    
    Поддерживает форматы:
    - Месяцы: '2025-март', '2024-янв', '2024-январь'
    - Недели: '2025/01', '2024/52' (год/номер через слеш), '2024-W01', '2024-W1', 
              '2024-нед01', '2024-нед1', '2024-н01'
    
    Args:
        period_str: Строка с периодом
        
    Returns:
        tuple: (year, period_number, type) где type: 0=месяц, 1=неделя
    """
    try:
        period_str = str(period_str).strip()
        
        # Сначала пытаемся распарсить как месяц
        match_month = re.match(r'(\d{4})[-_]?([а-яА-Я]+)', period_str.lower())
        if match_month:
            year = int(match_month.group(1))
            month_name = match_month.group(2)
            month = MONTHS_DICT.get(month_name, 0)
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


def parse_year_month(year_month_str):
    """Устаревшая функция, использует parse_period для обратной совместимости.
    
    Args:
        year_month_str: Строка с годом-месяцем
        
    Returns:
        tuple: (year, month)
    """
    result = parse_period(year_month_str)
    return (result[0], result[1])


def create_copy_button(text, button_label, key):
    """Создает кнопку для копирования текста в буфер обмена.
    
    Args:
        text: Текст для копирования
        button_label: Текст на кнопке
        key: Уникальный ключ для кнопки
    """
    # Очищаем key от специальных символов для использования в JavaScript
    safe_key = re.sub(r'[^a-zA-Z0-9_]', '_', str(key))
    
    # Экранируем текст для безопасной вставки в JavaScript
    # Используем JSON для правильного экранирования
    text_json = json.dumps(text)
    
    html = f"""
    <div data-testid="stButton" style="width: 100%; margin: 5px 0;">
        <button id="copy_btn_{safe_key}" onclick="copyToClipboard_{safe_key}()" style="
            width: 100%;
            padding: 12px 16px;
            background: #f8f9fa !important;
            color: #333 !important;
            border: 2px solid #e0e0e0 !important;
            border-radius: 8px !important;
            cursor: pointer !important;
            font-weight: 400 !important;
            font-size: 0.85rem !important;
            line-height: 1.3 !important;
            text-align: center !important;
            min-height: 50px !important;
            height: auto !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            white-space: normal !important;
            word-wrap: break-word !important;
            overflow-wrap: break-word !important;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05) !important;
            transition: all 0.3s ease !important;
            margin: 0 !important;
            box-sizing: border-box !important;
            position: relative !important;
        " onmouseover="if (!this.classList.contains('copied')) {{ this.style.transform='translateY(-2px)'; this.style.boxShadow='0 4px 8px rgba(0, 0, 0, 0.1)'; this.style.background='#ffffff'; this.style.borderColor='#d0d0d0'; }}" onmouseout="if (!this.classList.contains('copied')) {{ this.style.transform='translateY(0)'; this.style.boxShadow='0 2px 4px rgba(0, 0, 0, 0.05)'; this.style.background='#f8f9fa'; this.style.borderColor='#e0e0e0'; }}" onmousedown="if (!this.classList.contains('copied')) {{ this.style.transform='translateY(0)'; this.style.boxShadow='0 2px 4px rgba(0, 0, 0, 0.05)'; }}" onmouseup="if (!this.classList.contains('copied')) {{ this.style.transform='translateY(-2px)'; this.style.boxShadow='0 4px 8px rgba(0, 0, 0, 0.1)'; }}">
            <div style="display: flex; align-items: center; justify-content: center; width: 100%;">
                <p id="copy_btn_text_{safe_key}" style="margin: 0; padding: 0; font-size: 0.85rem; font-weight: 400; line-height: 1.3; word-wrap: break-word; overflow-wrap: break-word; white-space: normal;">{button_label}</p>
            </div>
        </button>
        <style>
            @keyframes pulse {{
                0% {{ transform: scale(1); }}
                50% {{ transform: scale(1.05); }}
                100% {{ transform: scale(1); }}
            }}
        </style>
    </div>
    <script>
        const textToCopy_{safe_key} = {text_json};
        
        function copyToClipboard_{safe_key}() {{
            const text = textToCopy_{safe_key};
            const button = document.getElementById('copy_btn_{safe_key}');
            const buttonText = document.getElementById('copy_btn_text_{safe_key}');
            const originalText = buttonText.innerHTML;
            
            // Функция для показа успешного копирования
            function showSuccess() {{
                // Изменяем внешний вид кнопки
                button.classList.add('copied');
                button.style.background = 'linear-gradient(135deg, #4CAF50 0%, #45a049 100%)';
                button.style.borderColor = '#4CAF50';
                button.style.color = 'white';
                button.style.transform = 'scale(0.98)';
                buttonText.innerHTML = '✓ Скопировано!';
                
                // Возвращаем исходное состояние через 2.5 секунды
                setTimeout(function() {{
                    button.classList.remove('copied');
                    button.style.background = '#f8f9fa';
                    button.style.borderColor = '#e0e0e0';
                    button.style.color = '#333';
                    button.style.transform = 'translateY(0)';
                    buttonText.innerHTML = originalText;
                }}, 2500);
            }}
            
            // Пробуем использовать современный API
            if (navigator.clipboard && navigator.clipboard.writeText) {{
                navigator.clipboard.writeText(text).then(function() {{
                    showSuccess();
                }}).catch(function(err) {{
                    console.error('Clipboard API error:', err);
                    // Fallback на старый метод
                    fallbackCopy_{safe_key}(text, showSuccess);
                }});
            }} else {{
                // Fallback для старых браузеров
                fallbackCopy_{safe_key}(text, showSuccess);
            }}
        }}
        
        function fallbackCopy_{safe_key}(text, successCallback) {{
            const textarea = document.createElement('textarea');
            textarea.value = text;
            textarea.style.position = 'fixed';
            textarea.style.left = '-999999px';
            textarea.style.top = '-999999px';
            textarea.style.opacity = '0';
            document.body.appendChild(textarea);
            textarea.focus();
            textarea.select();
            
            try {{
                const successful = document.execCommand('copy');
                if (successful) {{
                    successCallback();
                }} else {{
                    alert('Не удалось скопировать. Пожалуйста, скопируйте вручную.');
                }}
            }} catch(err) {{
                console.error('Copy command error:', err);
                alert('Ошибка копирования: ' + err);
            }} finally {{
                document.body.removeChild(textarea);
            }}
        }}
    </script>
    """
    components.html(html, height=70)


def detect_columns(df):
    """Автоматически определяет столбцы периода и клиента в DataFrame.
    
    Args:
        df: DataFrame с данными
        
    Returns:
        tuple: (year_month_col, client_col) или (None, None) если не найдены
    """
    year_month_col = None
    client_col = None
    
    # Ищем столбец с периодом (год-месяц или год-неделя)
    for col in df.columns:
        col_lower = str(col).lower()
        if 'год' in col_lower and ('месяц' in col_lower or 'неделя' in col_lower):
            year_month_col = col
            break
    
    # Ищем столбец с кодом клиента
    for col in df.columns:
        col_lower = str(col).lower()
        if 'код' in col_lower and 'клиент' in col_lower:
            client_col = col
            break
    
    return year_month_col, client_col


def get_sorted_periods(df, year_month_col):
    """Возвращает список периодов в том же порядке, что и matrix_builder (для согласованной когорты).
    
    Args:
        df: DataFrame с данными
        year_month_col: название столбца с периодом
        
    Returns:
        list: отсортированный список периодов
    """
    unique_periods = df[year_month_col].dropna().unique()
    periods_with_sort = [(p, parse_period(str(p).strip())) for p in unique_periods]
    valid_periods = [(p, parsed) for p, parsed in periods_with_sort if parsed != (0, 0, 0)]
    invalid_periods = [p for p, parsed in periods_with_sort if parsed == (0, 0, 0)]
    if valid_periods:
        valid_periods.sort(key=lambda x: (x[1][0], x[1][2], x[1][1]))
        sorted_periods = [p[0] for p in valid_periods]
        if invalid_periods:
            sorted_periods.extend(sorted(invalid_periods))
    else:
        sorted_periods = sorted([str(p) for p in unique_periods])
    return sorted_periods


def get_period_after_label(sorted_periods):
    """Возвращает подпись периода в родительном падеже для метрик: 'недели' или 'месяца'.
    
    Используется в подписях вида «после {label} когорты» в зависимости от того,
    построен ли анализ по неделям или по месяцам.
    
    Args:
        sorted_periods: отсортированный список периодов (из когортного анализа)
        
    Returns:
        str: 'недели' если периоды — недели, иначе 'месяца'
    """
    if not sorted_periods:
        return 'месяца'
    parsed = parse_period(str(sorted_periods[0]).strip())
    # type: 0 = месяц, 1 = неделя
    return 'недели' if parsed[2] == 1 else 'месяца'


def normalize_client_code(val):
    """Приводит код клиента к единому строковому виду для сравнения между файлами.
    
    Убирает пробелы; числа приводит к целому и строке, чтобы '196107' и '196107.0' совпадали.
    
    Args:
        val: значение из столбца (число или строка)
        
    Returns:
        str: нормализованный код клиента
    """
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ''
    s = str(val).strip().replace(' ', '')
    if not s:
        return ''
    try:
        return str(int(float(s)))
    except (ValueError, TypeError):
        return s


def normalize_period_for_compare(val):
    """Приводит период к каноническому виду для сравнения между файлами.
    
    Недели: '2025/1', '2025/01', 202501 -> '2025/01'. Месяцы возвращаются как str(val).strip().
    
    Args:
        val: значение периода из столбца Год-неделя/Год-месяц
        
    Returns:
        str: нормализованная строка периода
    """
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ''
    s = str(val).strip()
    if not s:
        return ''
    parsed = parse_period(s)
    if parsed == (0, 0, 0):
        return s
    year, num, ptype = parsed
    if ptype == 1:  # неделя
        return f"{year}/{num:02d}"
    if ptype == 0:  # месяц
        return f"{year}-{num:02d}"
    return s

