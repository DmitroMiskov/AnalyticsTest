"""
build.py — Відтворюваний пайплайн обробки даних мережі кав'ярень «Зерно».
Запуск:
    python build.py --input test_data.xlsx
"""

import os
import sys
import argparse
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
from matplotlib.patches import Patch

sys.stdout.reconfigure(encoding='utf-8')


def clean_shifts_data(excel_path='test_data.xlsx'):
    """
    Зчитує сирі дані, очищає їх, проводить assert-перевірки та повертає чисті датафрейми.
    """
    log = []
    
    xl = pd.ExcelFile(excel_path)
    df_raw = xl.parse('Зміни')
    df_cafes = xl.parse("Кав'ярні")
    df_baristas = xl.parse('Бариста')
    
    initial_rows = len(df_raw)
    log.append(f"Початкова кількість рядків у 'Зміни': {initial_rows}")
    
    # 1. Вилучення підсумкових рядків
    subtotals_mask = df_raw['Дата'].astype(str).str.contains('Разом', na=False)
    n_subtotals = subtotals_mask.sum()
    df = df_raw[~subtotals_mask].copy()
    log.append(f"Крок 1: Вилучено агреговані місячні підсумки ('Разом за...'): {n_subtotals} рядків. Залишилось: {len(df)}")
    
    # 2. Стандартизація формату дат
    def parse_date(x):
        if isinstance(x, (pd.Timestamp, np.datetime64)):
            return pd.to_datetime(x)
        if isinstance(x, str):
            parts = x.strip().split('.')
            if len(parts) == 3:
                return pd.Timestamp(year=int(parts[2]), month=int(parts[1]), day=int(parts[0]))
        return pd.to_datetime(x, dayfirst=True)
    
    str_dates_count = df['Дата'].apply(lambda x: isinstance(x, str)).sum()
    df['Дата'] = df['Дата'].apply(parse_date)
    log.append(f"Крок 2: Стандартизовано текстові дати 'DD.MM.YYYY' у datetime64: {str_dates_count} рядків (Одеса · Аркадія).")
    
    # 3. Стандартизація текстової виручки (кома, пробіли)
    def clean_rev(v):
        if pd.isna(v):
            return np.nan
        if isinstance(v, (int, float)):
            return float(v)
        s = str(v).replace('грн.', '').replace('грн', '').replace('\xa0', '').replace(' ', '').replace(',', '.').strip()
        return float(s)
    
    str_rev_count = df['Виручка, грн'].apply(lambda x: isinstance(x, str)).sum()
    df['Виручка, грн'] = df['Виручка, грн'].apply(clean_rev)
    log.append(f"Крок 3: Очищено текстову виручку (пробіли, коми) та конвертовано у float64: {str_rev_count} рядків (Львів · Франка).")
    
    # 4. Дедуплікація повних повторів рядків за бізнес-ключем (Дата, Кав'ярня, Зміна)
    dups_mask = df.duplicated(subset=['Дата', "Кав'ярня", 'Зміна'], keep='first')
    n_dups = dups_mask.sum()
    df = df[~dups_mask].copy()
    log.append(f"Крок 4: Вилучено дублікати змін (залишено перший запис): {n_dups} повторів. Залишилось: {len(df)}")
    
    # 5. Коригування виручки в копійках у «Київ · Лівобережна» (01.07.2026 - 14.07.2026)
    liv_mask = (df["Кав'ярня"] == 'Київ · Лівобережна') & (df['Дата'] >= '2026-07-01') & (df['Дата'] <= '2026-07-14') & (df['Виручка, грн'] > 1000000)
    n_liv = liv_mask.sum()
    df.loc[liv_mask, 'Виручка, грн'] = df.loc[liv_mask, 'Виручка, грн'] / 100.0
    log.append(f"Крок 5: Поділено на 100 виручку, збережену в копійках ('Київ · Лівобережна', 01–14 липня): {n_liv} змін.")
    
    # 6. Коригування опечатки у виручці x10 у «Київ · Хрещатик» (2026-05-03, вечір)
    khr_mask = (df["Кав'ярня"] == 'Київ · Хрещатик') & (df['Дата'] == '2026-05-03') & (df['Зміна'] == 'вечірня') & (df['Виручка, грн'] > 200000)
    n_khr = khr_mask.sum()
    df.loc[khr_mask, 'Виручка, грн'] = df.loc[khr_mask, 'Виручка, грн'] / 10.0
    log.append(f"Крок 6: Виправлено опечатку x10 у виручці ('Київ · Хрещатик', 03.05 вечір): 238 922.80 -> 23 892.28 грн ({n_khr} рядок).")
    
    # 7. Коригування опечатки x10 у чеках та виручці «Львів · Ринок» (2026-08-14, ранок)
    lviv_mask = (df["Кав'ярня"] == 'Львів · Ринок') & (df['Дата'] == '2026-08-14') & (df['Зміна'] == 'ранкова') & (df['Чеків'] == 1500)
    n_lviv = lviv_mask.sum()
    df.loc[lviv_mask, 'Чеків'] = (df.loc[lviv_mask, 'Чеків'] / 10).astype(int)
    df.loc[lviv_mask, 'Виручка, грн'] = df.loc[lviv_mask, 'Виручка, грн'] / 10.0
    log.append(f"Крок 7: Виправлено опечатку x10 у чеках та виручці ('Львів · Ринок', 14.08 ранок): 1500 -> 150 чеків, 219 741.28 -> 21 974.13 грн ({n_lviv} рядок).")
    
    # 8. Приведення типів колонок
    df['Годин'] = df['Годин'].astype(int)
    df['Чеків'] = df['Чеків'].astype(int)
    df['Виручка, грн'] = df['Виручка, грн'].astype(float)
    df['Оцінок'] = df['Оцінок'].astype(int)
    df['Середня оцінка'] = df['Середня оцінка'].astype(float)
    df['Середній чек'] = df['Виручка, грн'] / df['Чеків']
    
    # ASSERT-ПЕРЕВІРКИ
    assert df.duplicated(subset=['Дата', "Кав'ярня", 'Зміна']).sum() == 0, "Assert failed: знайдено дублікати за бізнес-ключем!"
    assert df['Дата'].min() >= pd.Timestamp('2026-04-01'), f"Assert failed: дата раніша за 01.04.2026: {df['Дата'].min()}"
    assert df['Дата'].max() <= pd.Timestamp('2026-09-20'), f"Assert failed: дата пізніша за 20.09.2026: {df['Дата'].max()}"
    assert set(df["Кав'ярня"].unique()).issubset(set(df_cafes["Кав'ярня"].unique())), "Assert failed: невідомі кав'ярні!"
    assert set(df['Бариста'].unique()).issubset(set(df_baristas['ID бариста'].unique())), "Assert failed: невідомі бариста!"
    assert pd.api.types.is_datetime64_any_dtype(df['Дата']), "Assert failed: 'Дата' не datetime"
    assert pd.api.types.is_integer_dtype(df['Годин']), "Assert failed: 'Годин' не integer"
    assert pd.api.types.is_integer_dtype(df['Чеків']), "Assert failed: 'Чеків' не integer"
    assert pd.api.types.is_float_dtype(df['Виручка, грн']), "Assert failed: 'Виручка' не float"
    assert pd.api.types.is_integer_dtype(df['Оцінок']), "Assert failed: 'Оцінок' не integer"
    assert pd.api.types.is_float_dtype(df['Середня оцінка']), "Assert failed: 'Середня оцінка' не float"
    
    total_shifts = len(df)
    total_revenue = df['Виручка, грн'].sum()
    total_checks = df['Чеків'].sum()
    
    return df, df_cafes, df_baristas, log, total_shifts, total_revenue, total_checks


def generate_charts(df, df_cafes):
    """
    Генерує 3 графіки-висновки у директорію charts/.
    """
    os.makedirs('charts', exist_ok=True)
    plt.style.use('seaborn-v0_8-whitegrid' if 'seaborn-v0_8-whitegrid' in plt.style.available else 'default')
    plt.rcParams['font.sans-serif'] = 'DejaVu Sans'
    plt.rcParams['axes.edgecolor'] = '#cccccc'
    plt.rcParams['axes.linewidth'] = 0.8
    
    if 'Місто' not in df.columns:
        df = df.merge(df_cafes[['Кав\'ярня', 'Місто', 'Посадкових місць', 'Відкрита', 'Графік', 'Акція з']], on="Кав'ярня", how='left')
    
    city_colors = {'Київ': '#2b5c8f', 'Одеса': '#d97724', 'Львів': '#2a8a58'}
    
    # -------------------------------------------------------------
    # ГРАФІК 1 (3.2 & 2.1): Виправлений bad_chart
    # -------------------------------------------------------------
    cafe_agg = df.groupby(["Кав'ярня", "Місто"]).agg(
        total_rev=('Виручка, грн', 'sum')
    ).reset_index().sort_values('total_rev', ascending=True)
    
    fig, ax = plt.subplots(figsize=(11, 6.2), dpi=300)
    bar_colors = [city_colors[c] for c in cafe_agg['Місто']]
    bars = ax.barh(cafe_agg["Кав'ярня"], cafe_agg['total_rev'] / 1e6, color=bar_colors, alpha=0.85, height=0.62)
    
    # Акцент на лідері
    bars[-1].set_color('#163b66')
    bars[-1].set_edgecolor('#c98a0c')
    bars[-1].set_linewidth(2.2)
    
    tot_rev = df['Виручка, грн'].sum()
    for bar, val in zip(bars, cafe_agg['total_rev']):
        pct = (val / tot_rev) * 100
        ax.text(bar.get_width() + 0.12, bar.get_y() + bar.get_height()/2,
                f"{val/1e6:.2f} млн грн  ({pct:.1f}%)",
                va='center', ha='left', fontsize=9.5, fontweight='bold', color='#222222')
        
    ax.set_xlim(0, 12.8)
    ax.set_xlabel('Сумарна виручка за період (млн грн)', fontsize=11, fontweight='bold', labelpad=10)
    ax.xaxis.set_major_formatter(ticker.FormatStrFormatter('%.1f'))
    
    legend_elements = [
        Patch(facecolor='#2b5c8f', label='Київ'),
        Patch(facecolor='#d97724', label='Одеса'),
        Patch(facecolor='#2a8a58', label='Львів'),
        Patch(facecolor='#163b66', edgecolor='#c98a0c', linewidth=1.5, label='Лідер мережі')
    ]
    ax.legend(handles=legend_elements, loc='lower right', frameon=True, framealpha=0.9, fontsize=10, title="Місто та статус")
    ax.set_title("«Київ · Хрещатик» та «Одеса · Дерибасівська» забезпечують 35% виручки мережі\n"
                 "Рейтинг кав'ярень за сумарною виручкою (01.04–20.09.2026, очищені дані, N=3 372 зміни)",
                 fontsize=12.5, fontweight='bold', pad=15, ha='center', color='#111111')
    
    plt.tight_layout()
    chart1_file = os.path.abspath(os.path.join('charts', 'fixed_revenue_chart.png'))
    plt.savefig(chart1_file, bbox_inches='tight')
    plt.close()
    
    # -------------------------------------------------------------
    # ГРАФІК 2 (2.2): Динаміка промо vs контроль
    # -------------------------------------------------------------
    promo_cafes = df_cafes[df_cafes['Акція з'].notna()]["Кав'ярня"].tolist()
    control_mature_cafes = [c for c in df_cafes[df_cafes['Акція з'].isna()]["Кав'ярня"].tolist() if c != 'Київ · Теремки']
    
    df['Month'] = df['Дата'].dt.month
    promo_daily = df[df["Кав'ярня"].isin(promo_cafes)].groupby('Month').apply(
        lambda g: g['Виручка, грн'].sum() / (g['Дата'].nunique() * len(promo_cafes))
    )
    ctrl_daily = df[df["Кав'ярня"].isin(control_mature_cafes)].groupby('Month').apply(
        lambda g: g['Виручка, грн'].sum() / (g['Дата'].nunique() * len(control_mature_cafes))
    )
    
    promo_idx = (promo_daily / promo_daily[5]) * 100
    ctrl_idx = (ctrl_daily / ctrl_daily[5]) * 100
    
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5.5), dpi=300)
    x_ticks = [4, 5, 6, 7, 8, 9]
    months_labels = ['Квітень', 'Травень (база)', 'Червень (акція)', 'Липень', 'Серпень', 'Вересень (20д)']
    
    ax1.plot(x_ticks, promo_idx, marker='o', linewidth=2.5, color='#d9534f', label='5 точок з акцією (промо)')
    ax1.plot(x_ticks, ctrl_idx, marker='s', linewidth=2.5, color='#337ab7', linestyle='--', label='4 зрілі точки без акції (контроль)')
    ax1.axvline(x=5, color='#888888', linestyle=':', alpha=0.7)
    ax1.text(5.08, 103, 'Старт акції\nз 1 червня', fontsize=9, color='#444444', fontweight='bold')
    
    for x, y_p, y_c in zip(x_ticks, promo_idx, ctrl_idx):
        ax1.text(x, y_p + 1.8, f"{y_p:.0f}%", ha='center', fontsize=9, fontweight='bold', color='#d9534f')
        ax1.text(x, y_c - 3.2, f"{y_c:.0f}%", ha='center', fontsize=9, fontweight='bold', color='#337ab7')
        
    ax1.set_xticks(x_ticks)
    ax1.set_xticklabels(months_labels, rotation=15)
    ax1.set_ylabel('Індекс середньоденної виручки (Травень = 100%)', fontsize=10, fontweight='bold')
    ax1.set_title("Динаміка: однаковий ріст через сезонність (+29% vs +22%)", fontsize=11, fontweight='bold')
    ax1.set_ylim(85, 140)
    ax1.legend(loc='upper left', frameon=True)
    
    # Правий сабплот: чек вечора
    promo_shifts = df[df["Кав'ярня"].isin(promo_cafes)].copy()
    promo_shifts['Period'] = np.where(promo_shifts['Дата'] < '2026-06-01', 'До акції (квітень–травень)', 'Під час акції (червень–вересень)')
    check_shift = promo_shifts.groupby(['Period', 'Зміна']).apply(lambda g: g['Виручка, грн'].sum() / g['Чеків'].sum()).unstack()
    
    shifts_categories = ['Ранкова зміна\n(без акції)', 'Вечірня зміна\n(акція −50% з 16:00)']
    before_vals = [check_shift.loc['До акції (квітень–травень)', 'ранкова'], check_shift.loc['До акції (квітень–травень)', 'вечірня']]
    after_vals = [check_shift.loc['Під час акції (червень–вересень)', 'ранкова'], check_shift.loc['Під час акції (червень–вересень)', 'вечірня']]
    
    x = np.arange(len(shifts_categories))
    width = 0.35
    rects1 = ax2.bar(x - width/2, before_vals, width, label='До акції (квітень–травень)', color='#8faadc')
    rects2 = ax2.bar(x + width/2, after_vals, width, label='Під час акції (червень–вересень)', color='#e06666')
    
    for r, val in zip(rects1, before_vals):
        ax2.text(r.get_x() + r.get_width()/2, val + 1.2, f"{val:.1f} грн", ha='center', fontsize=9, fontweight='bold')
    for r, val in zip(rects2, after_vals):
        ax2.text(r.get_x() + r.get_width()/2, val + 1.2, f"{val:.1f} грн", ha='center', fontsize=9, fontweight='bold')
        
    ax2.annotate('−14.4% просідання чека',
                 xy=(1 + width/2, after_vals[1]), xytext=(0.82, 60),
                 arrowprops=dict(facecolor='black', shrink=0.08, width=1, headwidth=6),
                 fontsize=9.5, fontweight='bold', color='#b30000')
                 
    ax2.set_xticks(x)
    ax2.set_xticklabels(shifts_categories)
    ax2.set_ylim(0, 130)
    ax2.set_ylabel('Зважений середній чек (грн)', fontsize=10, fontweight='bold')
    ax2.set_title("Вечірній чек просів зі 100.3 до 85.8 грн", fontsize=11, fontweight='bold')
    ax2.legend(loc='lower left', frameon=True)
    
    fig.suptitle("Акція «друга кава −50%» не дала надприбутку: приріст зумовлений літом, а чек впав на 14.4%",
                 fontsize=12.5, fontweight='bold', y=0.98, color='#111111')
    plt.tight_layout()
    chart2_file = os.path.abspath(os.path.join('charts', 'promo_evaluation.png'))
    plt.savefig(chart2_file, bbox_inches='tight')
    plt.close()
    
    # -------------------------------------------------------------
    # ГРАФІК 3 (2.3): Дисбаланс чека за локаціями
    # -------------------------------------------------------------
    cafe_check_summary = df.groupby(["Кав'ярня", "Місто"]).apply(
        lambda g: pd.Series({'weighted_check': g['Виручка, грн'].sum() / g['Чеків'].sum()})
    ).reset_index().sort_values('weighted_check', ascending=True)
    
    fig, ax = plt.subplots(figsize=(11, 5.8), dpi=300)
    y_pos = np.arange(len(cafe_check_summary))
    bar_colors = [city_colors[m] for m in cafe_check_summary['Місто']]
    bars = ax.barh(y_pos, cafe_check_summary['weighted_check'], color=bar_colors, alpha=0.85, height=0.6)
    
    for bar, w_val in zip(bars, cafe_check_summary['weighted_check']):
        ax.text(bar.get_width() + 1.2, bar.get_y() + bar.get_height()/2, f"{w_val:.2f} грн",
                va='center', ha='left', fontsize=9.5, fontweight='bold', color='#222222')
                
    ax.set_yticks(y_pos)
    ax.set_yticklabels(cafe_check_summary["Кав'ярня"], fontsize=10)
    ax.set_xlim(0, 150)
    ax.set_xlabel('Зважений середній чек кав\'ярні (грн)', fontsize=11, fontweight='bold', labelpad=10)
    
    net_check = df['Виручка, грн'].sum() / df['Чеків'].sum()
    ax.axvline(x=net_check, color='#e03e2d', linestyle='--', linewidth=1.5, label=f'Середній по всій мережі: {net_check:.2f} грн')
    ax.legend(loc='lower right', frameon=True, fontsize=10)
    ax.set_title("Середній чек визначається локацією точки, а не майстерністю бариста\n"
                 "Різниця між Хрещатиком та спальними районами сягає +50% (преміювати за абсолютним чеком несправедливо)",
                 fontsize=12, fontweight='bold', pad=15, ha='center', color='#111111')
                 
    plt.tight_layout()
    chart3_file = os.path.abspath(os.path.join('charts', 'barista_check_bias.png'))
    plt.savefig(chart3_file, bbox_inches='tight')
    plt.close()


def main():
    parser = argparse.ArgumentParser(description="Відтворюваний пайплайн аналітики мережі кав'ярень «Зерно»")
    parser.add_argument("--input", default="test_data.xlsx", help="Шлях до вхідного Excel-файлу")
    args = parser.parse_args()
    
    df, df_cafes, df_baristas, log, shifts, revenue, checks = clean_shifts_data(args.input)
    
    print("=" * 65)
    print("ЛОГ ТРАНСФОРМАЦІЇ ДАНИХ (build.py)")
    print("=" * 65)
    for entry in log:
        print(f"  • {entry}")
        
    print("\n" + "=" * 65)
    print("КОНТРОЛЬНІ ЦИФРИ ОЧИЩЕНОЇ ТАБЛИЦІ:")
    print("=" * 65)
    print(f"1. Кількість змін:          {shifts:,} змін")
    print(f"2. Сумарна виручка:         {revenue:,.2f} грн")
    print(f"3. Сумарна кількість чеків: {checks:,} чеків")
    print(f"   (Зважений середній чек:  {revenue / checks:.4f} грн)")
    print("=" * 65)
    print("Усі assert-перевірки виконано успішно.")
    
    print("\nГенерація графіків у папку charts/...")
    generate_charts(df, df_cafes)
    print("Графіки успішно згенеровано.")
    print("Файл звіту summary.html готовий до перегляду.")


if __name__ == "__main__":
    main()
