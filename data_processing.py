"""
Модуль для обработки данных и работы с клиентами
"""
import pandas as pd
from utils import get_sorted_periods


def get_cohort_clients(df, year_month_col, client_col, cohort_period, target_period, period_clients_cache=None, client_cohorts_cache=None, sorted_periods=None):
    """Получает коды клиентов из когорты (период первой покупки), которые были в целевом периоде.
    
    Args:
        df: DataFrame с данными
        year_month_col: название столбца с периодом
        client_col: название столбца с кодом клиента
        cohort_period: период когорты (первая покупка)
        target_period: целевой период
        period_clients_cache: кэш период -> множество клиентов
        client_cohorts_cache: кэш клиент -> период когорты (первая покупка)
        sorted_periods: отсортированный список периодов (нужен при client_cohorts_cache=None)
        
    Returns:
        list: отсортированный список кодов клиентов
    """
    if client_cohorts_cache is None:
        if sorted_periods is None:
            sorted_periods = get_sorted_periods(df, year_month_col)
        client_cohorts_cache = get_client_cohorts(df, year_month_col, client_col, sorted_periods)
    clients_in_cohort = {c for c, first in client_cohorts_cache.items() if first == cohort_period}
    if period_clients_cache:
        clients_in_period = period_clients_cache.get(target_period, set())
    else:
        clients_in_period = set(df[df[year_month_col] == target_period][client_col].dropna().unique())
    return sorted(list(clients_in_cohort & clients_in_period))


def get_accumulation_clients(df, year_month_col, client_col, sorted_periods, cohort_period, target_period, period_clients_cache=None, client_cohorts_cache=None):
    """Получает накопленные коды клиентов из когорты (период первой покупки) до целевого периода включительно.
    
    Args:
        df: DataFrame с данными
        year_month_col: название столбца с периодом
        client_col: название столбца с кодом клиента
        sorted_periods: отсортированный список периодов
        cohort_period: период когорты (первая покупка)
        target_period: целевой период
        period_clients_cache: кэш период -> множество клиентов
        client_cohorts_cache: кэш клиент -> период когорты (первая покупка)
        
    Returns:
        list: отсортированный список кодов клиентов
    """
    period_indices = {period: idx for idx, period in enumerate(sorted_periods)}
    cohort_idx = period_indices.get(cohort_period, -1)
    target_idx = period_indices.get(target_period, -1)
    
    if cohort_idx < 0 or target_idx < 0 or target_idx <= cohort_idx:
        return []
    
    if client_cohorts_cache is None:
        client_cohorts_cache = get_client_cohorts(df, year_month_col, client_col, sorted_periods)
    cohort_clients = {c for c, first in client_cohorts_cache.items() if first == cohort_period}
    
    returned_clients = set()
    for period in sorted_periods[cohort_idx + 1:target_idx + 1]:
        if period_clients_cache:
            period_clients = period_clients_cache.get(period, set())
        else:
            period_clients = set(df[df[year_month_col] == period][client_col].dropna().unique())
        returned_clients.update(cohort_clients & period_clients)
    
    return sorted(list(returned_clients))


def get_client_cohorts(df, year_month_col, client_col, sorted_periods):
    """Определяет когорту для каждого клиента (первый период появления по порядку sorted_periods).
    
    Args:
        df: DataFrame с данными
        year_month_col: название столбца с периодом
        client_col: название столбца с кодом клиента
        sorted_periods: отсортированный список периодов
        
    Returns:
        dict: словарь {client: cohort_period}
    """
    period_indices = {period: idx for idx, period in enumerate(sorted_periods)}
    df_filtered = df[[year_month_col, client_col]].dropna()
    client_cohorts = {}
    for client, group in df_filtered.groupby(client_col):
        periods = group[year_month_col].dropna().unique()
        valid = [p for p in periods if p in period_indices]
        if valid:
            first_period = min(valid, key=lambda p: period_indices[p])
            client_cohorts[client] = first_period
    return client_cohorts


def get_churn_clients(df, year_month_col, client_col, sorted_periods, cohort_period, period_clients_cache=None, client_cohorts_cache=None):
    """Получает коды клиентов оттока из когорты.
    
    Отток = клиенты когорты, которые не вернулись ни разу после периода когорты.
    Когорта определяется как первый период появления клиента.
    
    Args:
        df: DataFrame с данными
        year_month_col: название столбца с периодом
        client_col: название столбца с кодом клиента
        sorted_periods: отсортированный список периодов
        cohort_period: период когорты
        period_clients_cache: кэш словарь период -> множество клиентов
        client_cohorts_cache: кэш словарь клиент -> период когорты
        
    Returns:
        list: отсортированный список кодов клиентов оттока
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


def get_inflow_clients(df, year_month_col, client_col, sorted_periods, cohort_period, target_period, period_clients_cache=None, client_cohorts_cache=None):
    """Получает коды клиентов из когорты (период первой покупки), которые вернулись именно в целевом периоде (новый приток).
    
    Args:
        df: DataFrame с данными
        year_month_col: название столбца с периодом
        client_col: название столбца с кодом клиента
        sorted_periods: отсортированный список периодов
        cohort_period: период когорты (первая покупка)
        target_period: целевой период
        period_clients_cache: кэш период -> множество клиентов
        client_cohorts_cache: кэш клиент -> период когорты (первая покупка)
        
    Returns:
        list: отсортированный список кодов клиентов
    """
    period_indices = {period: idx for idx, period in enumerate(sorted_periods)}
    cohort_idx = period_indices.get(cohort_period, -1)
    target_idx = period_indices.get(target_period, -1)
    
    if cohort_idx < 0 or target_idx < 0 or target_idx <= cohort_idx:
        return []
    
    if client_cohorts_cache is None:
        client_cohorts_cache = get_client_cohorts(df, year_month_col, client_col, sorted_periods)
    cohort_clients = {c for c, first in client_cohorts_cache.items() if first == cohort_period}
    
    if period_clients_cache:
        target_period_clients = period_clients_cache.get(target_period, set())
    else:
        target_period_clients = set(df[df[year_month_col] == target_period][client_col].dropna().unique())
    returned_in_target = cohort_clients & target_period_clients
    
    if target_idx == cohort_idx + 1:
        return sorted(list(returned_in_target))
    
    prev_periods_clients = set()
    for period in sorted_periods[cohort_idx + 1:target_idx]:
        if period_clients_cache:
            period_clients = period_clients_cache.get(period, set())
        else:
            period_clients = set(df[df[year_month_col] == period][client_col].dropna().unique())
        prev_periods_clients.update(cohort_clients & period_clients)
    
    new_returns = returned_in_target - prev_periods_clients
    return sorted(list(new_returns))


def create_period_clients_cache(df, year_month_col, client_col, sorted_periods):
    """Создает кэш период -> множество клиентов для оптимизации.
    
    Args:
        df: DataFrame с данными
        year_month_col: название столбца с периодом
        client_col: название столбца с кодом клиента
        sorted_periods: отсортированный список периодов
        
    Returns:
        dict: словарь период -> множество клиентов
    """
    period_clients_cache = {}
    for period in sorted_periods:
        period_data = df[df[year_month_col] == period]
        period_clients_cache[period] = set(period_data[client_col].dropna().unique())
    return period_clients_cache


def build_churn_table(df, year_month_col, client_col, sorted_periods, cohort_matrix, 
                       accumulation_matrix, accumulation_percent_matrix, 
                       client_cohorts_cache=None, period_clients_cache=None):
    """Строит таблицу оттока клиентов для всех когорт.
    
    Когорта = период первой покупки клиента. Размер когорты и отток считаются по этой логике.
    
    Args:
        df: DataFrame с данными
        year_month_col: название столбца с периодом
        client_col: название столбца с кодом клиента
        sorted_periods: отсортированный список периодов
        cohort_matrix: матрица когорт
        accumulation_matrix: матрица накопления
        accumulation_percent_matrix: матрица накопления в процентах
        client_cohorts_cache: кэш словарь клиент -> период когорты
        period_clients_cache: кэш словарь период -> множество клиентов
        
    Returns:
        pd.DataFrame: таблица оттока
    """
    churn_data = []
    
    # Оптимизация: создаём period_indices один раз вне цикла
    period_indices = {period: idx for idx, period in enumerate(sorted_periods)}
    last_period = sorted_periods[-1]
    last_period_idx = period_indices[last_period]
    
    for cohort_period in sorted_periods:
        cohort = cohort_period
        cohort_size = cohort_matrix.loc[cohort_period, cohort_period]
        cohort_idx = period_indices[cohort_period]
        is_last_cohort = (cohort_idx == last_period_idx)
        
        if is_last_cohort:
            # Для последней когорты нет периодов наблюдения после — не считаем возврат и отток
            churn_data.append({
                'Когорта': cohort,
                'Кол-во клиентов когорты': int(cohort_size),
                'Накопительное кол-во возврата': '-',
                'Накопительный % возврата': '-',
                'Отток кол-во': '-',
                'Отток %': '-'
            })
            continue
        
        total_returned = accumulation_matrix.loc[cohort_period, last_period]
        if cohort_size > 0:
            total_returned_percent = (total_returned / cohort_size) * 100
        else:
            total_returned_percent = 0
        churn_count = int(cohort_size - total_returned)
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

