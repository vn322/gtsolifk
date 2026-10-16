"""
Модуль статистического анализа: шаблон, анализ с размерами эффекта, HTML-отчёт с пояснениями.
"""
import os
import base64
import io
import warnings
from datetime import datetime

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats
import statsmodels.api as sm
from statsmodels.formula.api import ols

warnings.filterwarnings('ignore')

# ═════════════════════════════════════════════════════════════════════════════
# РАСЧЁТ РАЗМЕРОВ ЭФФЕКТА
# ═════════════════════════════════════════════════════════════════════════════
def cohens_d_paired(x, y):
    """Коэн d для парных выборок (внутригрупповой)"""
    diff = x - y
    sd = np.std(diff, ddof=1)
    return np.mean(diff) / sd if sd > 0 else 0.0

def cohens_d_independent(x, y):
    """Коэн d для независимых выборок (межгрупповой)"""
    n1, n2 = len(x), len(y)
    var1, var2 = np.var(x, ddof=1), np.var(y, ddof=1)
    pooled_sd = np.sqrt(((n1 - 1) * var1 + (n2 - 1) * var2) / (n1 + n2 - 2))
    return (np.mean(x) - np.mean(y)) / pooled_sd if pooled_sd > 0 else 0.0

def cliffs_delta(x, y):
    """Дельта Клиффа для непараметрических межгрупповых сравнений"""
    greater = sum(1 for i in x for j in y if i > j)
    less = sum(1 for i in x for j in y if i < j)
    n = len(x) * len(y)
    return (greater - less) / n if n > 0 else 0.0

def interpret_effect_size(val, test_type):
    """Интерпретация размера эффекта"""
    if test_type == 'cohen':
        if abs(val) >= 0.8: return f"{val:.2f} (большой)"
        if abs(val) >= 0.5: return f"{val:.2f} (средний)"
        if abs(val) >= 0.2: return f"{val:.2f} (малый)"
        return f"{val:.2f} (пренебрежимый)"
    elif test_type == 'rosenthal':
        if abs(val) >= 0.5: return f"{abs(val):.2f} (большой)"
        if abs(val) >= 0.3: return f"{abs(val):.2f} (средний)"
        if abs(val) >= 0.1: return f"{abs(val):.2f} (малый)"
        return f"{abs(val):.2f} (пренебрежимый)"
    elif test_type == 'cliff':
        if abs(val) >= 0.474: return f"{abs(val):.2f} (большой)"
        if abs(val) >= 0.33: return f"{abs(val):.2f} (средний)"
        if abs(val) >= 0.147: return f"{abs(val):.2f} (малый)"
        return f"{abs(val):.2f} (пренебрежимый)"
    return f"{val:.2f}"

# ═════════════════════════════════════════════════════════════════════════════
# ГЕНЕРАЦИЯ ШАБЛОНА
# ═════════════════════════════════════════════════════════════════════════════
def generate_template(output_path="шаблон_данных.xlsx"):
    """Генерирует чистый шаблон Excel: только заголовки и данные, без лишнего текста."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    sample_data = {
        'Спортсмен': ['Иванов А.', 'Петров Б.', 'Сидоров В.', 'Козлов Г.', 'Новиков Д.',
                      'Смирнов Е.', 'Кузнецов Ж.', 'Попов З.', 'Волков И.', 'Соколов К.'],
        'Группа': ['Контрольная']*5 + ['Экспериментальная']*5,
        'До': [10.5, 11.2, 9.8, 10.9, 11.5, 10.3, 11.0, 9.9, 10.7, 11.3],
        'После': [10.3, 11.0, 9.9, 10.8, 11.4, 9.5, 9.8, 8.9, 9.2, 9.7]
    }
    df = pd.DataFrame(sample_data)

    wb = Workbook()
    ws = wb.active
    ws.title = "Данные"

    # === ЗАГОЛОВКИ СТРОГО В ПЕРВОЙ СТРОКЕ (без инструкций над ними) ===
    headers = ['Спортсмен', 'Группа', 'До', 'После']
    header_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')
        cell.border = thin_border

    # === ДАННЫЕ НАЧИНАЮТСЯ СО ВТОРОЙ СТРОКИ ===
    for row_idx, row in enumerate(df.itertuples(index=False), 2):
        for col_idx, value in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center')
            if col_idx in [3, 4]:
                cell.number_format = '0.00'

    # Ширина столбцов
    for col, width in zip(['A', 'B', 'C', 'D'], [20, 22, 12, 12]):
        ws.column_dimensions[col].width = width

    # === ОТДЕЛЬНЫЙ ЛИСТ С ИНСТРУКЦИЕЙ (не мешает загрузке) ===
    ws2 = wb.create_sheet("Инструкция")
    instructions = [
        ["КАК ЗАПОЛНИТЬ ФАЙЛ"],
        [""],
        ["1. Перейдите на вкладку 'Данные'."],
        ["2. Столбец A 'Спортсмен' — ФИО или номер участника (текст)."],
        ["3. Столбец B 'Группа' — строго 'Контрольная' или 'Экспериментальная'."],
        ["4. Столбец C 'До' — показатель ДО воздействия (число через точку)."],
        ["5. Столбец D 'После' — показатель ПОСЛЕ воздействия (число через точку)."],
        [""],
        ["ВАЖНО:"],
        ["• Не удаляйте и не переименовывайте заголовки в первой строке."],
        ["• Минимум 5 человек в каждой группе."],
        ["• Не оставляйте пустых строк между данными."],
        ["• После заполнения сохраните файл и загрузите в программу."],
        [""],
        ["Вкладку 'Инструкция' можно удалить перед загрузкой — это не повлияет на анализ."]
    ]
    for row_idx, row in enumerate(instructions, 1):
        ws2.cell(row=row_idx, column=1, value=row[0])
    ws2['A1'].font = Font(bold=True, size=14)
    ws2.column_dimensions['A'].width = 80

    wb.save(output_path)
    return output_path

# ═════════════════════════════════════════════════════════════════════════════
# УТИЛИТЫ ОТЧЁТА
# ═════════════════════════════════════════════════════════════════════════════
def fig_to_base64(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white')
    buf.seek(0)
    img_str = base64.b64encode(buf.read()).decode('utf-8')
    buf.close()
    return f"data:image/png;base64,{img_str}"

def df_to_html(df, pvalue_cols=None):
    html = "<table><thead><tr>"
    for col in df.columns:
        html += f"<th>{col}</th>"
    html += "</tr></thead><tbody>"
    for _, row in df.iterrows():
        html += "<tr>"
        for col in df.columns:
            val = row[col]
            cell_class = ""
            if pvalue_cols and col in pvalue_cols:
                try:
                    p = float(val)
                    if p < 0.001: cell_class = 'sig-high'
                    elif p < 0.01: cell_class = 'sig-medium'
                    elif p < 0.05: cell_class = 'sig-low'
                    else: cell_class = 'sig-none'
                except (ValueError, TypeError):
                    pass
            display_val = f"{val:.4f}" if isinstance(val, float) and abs(val) < 1 else (f"{val:.2f}" if isinstance(val, float) else str(val))
            html += f'<td class="{cell_class}">{display_val}</td>'
        html += "</tr>"
    html += "</tbody></table>"
    return html

# ═════════════════════════════════════════════════════════════════════════════
# ОСНОВНАЯ ФУНКЦИЯ АНАЛИЗА
# ═════════════════════════════════════════════════════════════════════════════
def run_analysis(file_path, out_dir, log_callback=None):
    def log(msg):
        if log_callback: log_callback(msg)

    log(" Загрузка данных...")
    df = pd.read_excel(file_path)
    required = ['Спортсмен', 'Группа', 'До', 'После']
    missing = [c for c in required if c not in df.columns]
    if missing: raise ValueError(f"Отсутствуют столбцы: {missing}")

    df['Группа'] = df['Группа'].astype(str).str.strip()
    df['До'] = pd.to_numeric(df['До'], errors='coerce')
    df['После'] = pd.to_numeric(df['После'], errors='coerce')
    df = df.dropna(subset=['До', 'После'])

    ctrl_mask = df['Группа'].str.contains('Контроль', case=False, na=False)
    exp_mask = df['Группа'].str.contains('Эксперимент', case=False, na=False)
    if not ctrl_mask.any() or not exp_mask.any():
        raise ValueError("Группы должны содержать слова 'Контрольная' и 'Экспериментальная'")

    df_ctrl = df[ctrl_mask].copy()
    df_exp = df[exp_mask].copy()
    df_ctrl['Изменение'] = df_ctrl['После'] - df_ctrl['До']
    df_exp['Изменение'] = df_exp['После'] - df_exp['До']

    group_ctrl_name = df_ctrl['Группа'].iloc[0]
    group_exp_name = df_exp['Группа'].iloc[0]
    log(f"✓ Контрольная: {group_ctrl_name} (n={len(df_ctrl)})")
    log(f"✓ Экспериментальная: {group_exp_name} (n={len(df_exp)})")

    # 1. Описательная статистика
    log(" Расчёт описательной статистики...")
    def desc_stats(series, name):
        return {'Группа': name, 'n': len(series), 'M': series.mean(), 'SD': series.std(),
                'Md': series.median(), 'Q1': series.quantile(0.25), 'Q3': series.quantile(0.75),
                'IQR': series.quantile(0.75) - series.quantile(0.25), 'SE': series.std() / np.sqrt(len(series))}
    
    desc_rows = []
    for name, g in [('Контрольная (До)', df_ctrl['До']), ('Контрольная (После)', df_ctrl['После']),
                    ('Экспериментальная (До)', df_exp['До']), ('Экспериментальная (После)', df_exp['После'])]:
        desc_rows.append(desc_stats(g, name))
    desc_table = pd.DataFrame(desc_rows).round(3)

    # 2. Нормальность
    log(" Проверка нормальности (Шапиро-Уилк)...")
    norm_rows = []
    for name, g in [('Контрольная (До)', df_ctrl['До']), ('Контрольная (После)', df_ctrl['После']),
                    ('Экспериментальная (До)', df_exp['До']), ('Экспериментальная (После)', df_exp['После']),
                    ('Контрольная (Δ)', df_ctrl['Изменение']), ('Экспериментальная (Δ)', df_exp['Изменение'])]:
        if len(g) >= 3:
            stat, p = stats.shapiro(g)
            dp = stats.normaltest(g)[1] if len(g) >= 8 else np.nan
            norm_rows.append({'Группа': name, 'Shapiro W': stat, 'Shapiro p': p, 'DP p': dp})
    norm_table = pd.DataFrame(norm_rows).round(4)
    all_normal = (norm_table['Shapiro p'] >= 0.05).all()
    log(f"   Нормальность: {'✅ Да' if all_normal else '⚠ Нет'}")

    # 3. Внутригрупповые сравнения + Эффект
    log(" Внутригрупповые сравнения...")
    within_rows = []
    for name, g in [('Контрольная', df_ctrl), ('Экспериментальная', df_exp)]:
        if all_normal:
            stat, p = stats.ttest_rel(g['До'], g['После'])
            test = 't-тест парный'
            effect = cohens_d_paired(g['До'], g['После'])
            effect_str = interpret_effect_size(effect, 'cohen')
        else:
            res = stats.wilcoxon(g['До'], g['После'], method='approx')
            stat, p = res.statistic, res.pvalue
            test = 'Wilcoxon'
            z = res.statistic # В method='approx' statistic это Z
            r = abs(z) / np.sqrt(len(g))
            effect_str = interpret_effect_size(r, 'rosenthal')
        
        within_rows.append({'Группа': name, 'Тест': test, 'Статистика': round(stat, 3), 
                            'p-value': round(p, 4), 'Δ M±SD': f"{g['Изменение'].mean():.2f}±{g['Изменение'].std():.2f}",
                            'Размер эффекта': effect_str})
    within_table = pd.DataFrame(within_rows)

    # 4. Межгрупповые сравнения + Эффект
    log("🔄 Межгрупповые сравнения...")
    between_rows = []
    # Исходные
    if all_normal:
        stat, p = stats.ttest_ind(df_ctrl['До'], df_exp['До'])
        test, effect_str = 't-тест независ.', interpret_effect_size(cohens_d_independent(df_ctrl['До'], df_exp['До']), 'cohen')
    else:
        res = stats.mannwhitneyu(df_ctrl['До'], df_exp['До'], alternative='two-sided', method='asymptotic')
        stat, p = res.statistic, res.pvalue
        test = 'Mann-Whitney'
        z = res.statistic
        r = abs(z) / np.sqrt(len(df_ctrl) + len(df_exp))
        delta = cliffs_delta(df_ctrl['До'], df_exp['До'])
        effect_str = f"r={interpret_effect_size(r, 'rosenthal')}, Δ={interpret_effect_size(delta, 'cliff')}"
    between_rows.append({'Сравнение': 'До (между группами)', 'Тест': test, 'Статистика': round(stat, 3), 'p-value': round(p, 4), 'Размер эффекта': effect_str})

    # Изменения
    if all_normal:
        stat, p = stats.ttest_ind(df_ctrl['Изменение'], df_exp['Изменение'])
        test, effect_str = 't-тест независ.', interpret_effect_size(cohens_d_independent(df_ctrl['Изменение'], df_exp['Изменение']), 'cohen')
    else:
        res = stats.mannwhitneyu(df_ctrl['Изменение'], df_exp['Изменение'], alternative='two-sided', method='asymptotic')
        stat, p = res.statistic, res.pvalue
        test = 'Mann-Whitney'
        z = res.statistic
        r = abs(z) / np.sqrt(len(df_ctrl) + len(df_exp))
        delta = cliffs_delta(df_ctrl['Изменение'], df_exp['Изменение'])
        effect_str = f"r={interpret_effect_size(r, 'rosenthal')}, Δ={interpret_effect_size(delta, 'cliff')}"
    between_rows.append({'Сравнение': 'Изменения (между группами)', 'Тест': test, 'Статистика': round(stat, 3), 'p-value': round(p, 4), 'Размер эффекта': effect_str})
    between_table = pd.DataFrame(between_rows)

    # 5. ANOVA / ANCOVA
    log(" Построение математической модели...")
    anova_table, ancova_table = None, None
    try:
        long_rows = []
        for _, row in df_ctrl.iterrows():
            long_rows.append({'Спортсмен': row['Спортсмен'], 'Группа_f': 'Ctrl', 'Время': 'До', 'Значение': row['До']})
            long_rows.append({'Спортсмен': row['Спортсмен'], 'Группа_f': 'Ctrl', 'Время': 'После', 'Значение': row['После']})
        for _, row in df_exp.iterrows():
            long_rows.append({'Спортсмен': row['Спортсмен'], 'Группа_f': 'Exp', 'Время': 'До', 'Значение': row['До']})
            long_rows.append({'Спортсмен': row['Спортсмен'], 'Группа_f': 'Exp', 'Время': 'После', 'Значение': row['После']})
        long = pd.DataFrame(long_rows)
        model = ols('Значение ~ C(Группа_f) * C(Время)', data=long).fit()
        anova_table = sm.stats.anova_lm(model, typ=2).reset_index()
        anova_table = anova_table.rename(columns={'index': 'Эффект', 'PR(>F)': 'p-value'})[['Эффект', 'df', 'F', 'p-value']].round(4)

        df_ac = pd.concat([df_ctrl.assign(Группа=0), df_exp.assign(Группа=1)])
        model_ac = ols('После ~ Группа + До', data=df_ac).fit()
        ancova_table = pd.DataFrame({'Эффект': model_ac.model.exog_names, 'coef': model_ac.params.values, 'P>|t|': model_ac.pvalues.values}).round(4)
    except Exception as e:
        log(f"    ANOVA/ANCOVA: {e}")

    # 6. Графики
    log("🎨 Построение графиков...")
    os.makedirs(out_dir, exist_ok=True)
    figures = {}
    sns.set_style('whitegrid')
    plt.rcParams['font.family'] = ['DejaVu Sans', 'Arial']

    def save_fig(fig, name):
        path = os.path.join(out_dir, f"{name}.png")
        fig.savefig(path, dpi=150, bbox_inches='tight', facecolor='white')
        figures[name] = fig_to_base64(fig)
        plt.close(fig)

    df_plot = pd.concat([
        df_ctrl.assign(Группа=group_ctrl_name, Время='До')[['Группа', 'Время', 'До']].rename(columns={'До': 'Значение'}),
        df_ctrl.assign(Группа=group_ctrl_name, Время='После')[['Группа', 'Время', 'После']].rename(columns={'После': 'Значение'}),
        df_exp.assign(Группа=group_exp_name, Время='До')[['Группа', 'Время', 'До']].rename(columns={'До': 'Значение'}),
        df_exp.assign(Группа=group_exp_name, Время='После')[['Группа', 'Время', 'После']].rename(columns={'После': 'Значение'}),
    ])
    
    # 1. Violin
    fig, ax = plt.subplots(figsize=(8, 5))
    sns.violinplot(data=df_plot, x='Группа', y='Значение', hue='Время', split=False, ax=ax, palette='Set2')
    ax.set_title('Распределения До и После по группам')
    save_fig(fig, '01_violin_before_after')

    # 2. Boxplot
    fig, ax = plt.subplots(figsize=(7, 5))
    df_changes = pd.concat([df_ctrl.assign(Группа=group_ctrl_name)[['Группа', 'Изменение']],
                            df_exp.assign(Группа=group_exp_name)[['Группа', 'Изменение']]])
    sns.boxplot(data=df_changes, x='Группа', y='Изменение', ax=ax, palette='Set2')
    ax.axhline(0, color='red', linestyle='--')
    ax.set_title('Изменения показателей (После − До)')
    save_fig(fig, '02_boxplot_changes')

    # 3. Spaghetti
    fig, ax = plt.subplots(figsize=(9, 5))
    for _, row in df_ctrl.iterrows(): ax.plot(['До', 'После'], [row['До'], row['После']], 'o-', color='steelblue', alpha=0.4)
    for _, row in df_exp.iterrows(): ax.plot(['До', 'После'], [row['До'], row['После']], 'o-', color='coral', alpha=0.4)
    ax.set_title('Индивидуальные траектории'); ax.set_ylabel('Значение')
    save_fig(fig, '03_spaghetti_plot')

    # 4. Barplot
    fig, ax = plt.subplots(figsize=(8, 5))
    sns.barplot(data=df_plot, x='Группа', y='Значение', hue='Время', ax=ax, palette='Set2', errorbar=('pi', 50))
    ax.set_title('Медианы с межквартильным размахом (IQR)')
    save_fig(fig, '04_barplot_medians_iqr')

    # 5-6. QQ
    for name, data, fname in [('Контрольная', df_ctrl['Изменение'], '05_qqplot_control'), ('Экспериментальная', df_exp['Изменение'], '06_qqplot_experimental')]:
        fig = sm.qqplot(data, line='s')
        fig.axes[0].set_title(f'QQ-plot: {name}')
        save_fig(fig, fname)

    # 7-8. Histograms
    for name, data, fname in [('Контрольная', df_ctrl['Изменение'], '07_histogram_control'), ('Экспериментальная', df_exp['Изменение'], '08_histogram_experimental')]:
        fig, ax = plt.subplots(figsize=(7, 4))
        ax.hist(data, bins=8, color='steelblue', edgecolor='white')
        ax.set_title(f'Распределение изменений: {name}'); ax.set_xlabel('Δ значение')
        save_fig(fig, fname)

    # 9. Scatter
    fig, ax = plt.subplots(figsize=(7, 6))
    ax.scatter(df_ctrl['До'], df_ctrl['После'], label=group_ctrl_name, color='steelblue')
    ax.scatter(df_exp['До'], df_exp['После'], label=group_exp_name, color='coral')
    lims = [min(ax.get_xlim()[0], ax.get_ylim()[0]), max(ax.get_xlim()[1], ax.get_ylim()[1])]
    ax.plot(lims, lims, 'k--', alpha=0.5, label='y=x')
    ax.set_xlabel('До'); ax.set_ylabel('После'); ax.legend(); ax.set_title('Зависимость После от До')
    save_fig(fig, '09_scatter_before_after')

    # 10. Summary CI
    fig, ax = plt.subplots(figsize=(8, 5))
    def ci95(s):
        m, se = s.mean(), s.std() / np.sqrt(len(s))
        return (m - 1.96 * se, m + 1.96 * se)
    sns.pointplot(data=df_plot, x='Группа', y='Значение', hue='Время', ax=ax, palette='Set2', errorbar=ci95, capsize=0.1)
    ax.set_title('Средние значения с 95% доверительными интервалами')
    save_fig(fig, '10_summary_effects')

    # 7. Выводы
    conclusions = []
    p_ctrl = within_table.loc[within_table['Группа'] == 'Контрольная', 'p-value'].iloc[0]
    p_exp = within_table.loc[within_table['Группа'] == 'Экспериментальная', 'p-value'].iloc[0]
    
    conclusions.append(f"{'✅' if p_ctrl < 0.05 else '⚪'} Контрольная группа: изменения {'значимы' if p_ctrl < 0.05 else 'не значимы'} (p={p_ctrl:.4f}).")
    conclusions.append(f"{'✅' if p_exp < 0.05 else '⚪'} Экспериментальная группа: изменения {'значимы' if p_exp < 0.05 else 'не значимы'} (p={p_exp:.4f}).")
    
    p_between = between_table.iloc[1]['p-value']
    eff_between = between_table.iloc[1]['Размер эффекта']
    conclusions.append(f"{'🎯' if p_between < 0.05 else '⚪'} Различия в эффекте между группами: {'значимы' if p_between < 0.05 else 'не значимы'} (p={p_between:.4f}, эффект: {eff_between}).")
    
    if anova_table is not None:
        inter = anova_table[anova_table['Эффект'].str.contains('Группа.*Время|Время.*Группа', regex=True)]
        if not inter.empty and inter.iloc[0]['p-value'] < 0.05:
            conclusions.append(f"📐 Взаимодействие «Группа×Время» значимо (p={inter.iloc[0]['p-value']:.4f}) — методика имеет специфический эффект.")

    # 8. Генерация отчёта
    log(" Формирование HTML-отчёта...")
    base = os.path.splitext(os.path.basename(file_path))[0]
    report_path = os.path.join(out_dir, f"отчёт_{base}.html")
    
    results = {
        'file_name': os.path.basename(file_path), 'date': datetime.now().strftime("%d.%m.%Y %H:%M"),
        'n_total': len(df), 'group_ctrl': group_ctrl_name, 'n_ctrl': len(df_ctrl),
        'group_exp': group_exp_name, 'n_exp': len(df_exp), 'desc_table': desc_table,
        'norm_table': norm_table, 'all_normal': all_normal, 'within_table': within_table,
        'between_table': between_table, 'anova_table': anova_table, 'ancova_table': ancova_table,
        'figures': figures, 'conclusions': conclusions
    }
    generate_report(results, report_path)
    results['report_path'] = report_path
    log(f"✅ Отчёт сохранён: {report_path}")
    return results

# ═════════════════════════════════════════════════════════════════════════════
# ГЕНЕРАЦИЯ HTML-ОТЧЁТА С ПОДРОБНЫМИ ПОЯСНЕНИЯМИ
# ═════════════════════════════════════════════════════════════════════════════
def generate_report(results_dict, output_path="отчёт.html"):
    date_str = results_dict.get('date', datetime.now().strftime("%d.%m.%Y %H:%M"))
    css = """
    <style>
    @page { size: A4; margin: 1.5cm; }
    body { font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #2c3e50; max-width: 1100px; margin: 0 auto; padding: 20px; }
    h1 { color: #2c3e50; border-bottom: 3px solid #3498db; padding-bottom: 10px; }
    h2 { color: #34495e; margin-top: 35px; border-left: 5px solid #3498db; padding-left: 10px; }
    h3 { color: #7f8c8d; font-size: 1.1em; margin-top: 20px; }
    .meta { background: #ecf0f1; padding: 15px; border-radius: 8px; margin: 20px 0; }
    .meta p { margin: 5px 0; }
    table { border-collapse: collapse; width: 100%; margin: 15px 0; font-size: 13px; }
    th { background: #2c3e50; color: white; padding: 10px; text-align: center; }
    td { padding: 8px; border: 1px solid #bdc3c7; text-align: center; }
    tr:nth-child(even) { background: #f8f9fa; }
    .sig-high { background: #27ae60 !important; color: white; font-weight: bold; }
    .sig-medium { background: #2ecc71 !important; color: white; font-weight: bold; }
    .sig-low { background: #f1c40f !important; color: black; font-weight: bold; }
    .sig-none { background: #e74c3c !important; color: white; }
    .figure { text-align: center; margin: 25px 0; page-break-inside: avoid; }
    .figure img { max-width: 100%; border: 1px solid #bdc3c7; border-radius: 4px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
    .figure-caption { font-style: italic; color: #7f8c8d; margin-top: 8px; font-size: 14px; }
    .conclusion { background: #e8f6f3; border-left: 5px solid #16a085; padding: 15px; margin: 10px 0; border-radius: 4px; }
    .warning { background: #fef5e7; border-left: 5px solid #f39c12; padding: 15px; margin: 10px 0; border-radius: 4px; }
    .explanation { background: #f4f6f7; padding: 12px 15px; border-radius: 4px; font-size: 14px; color: #444; margin: 10px 0 15px 0; border-left: 3px solid #95a5a6; }
    .explanation b { color: #2c3e50; }
    .step-intro { background: #eaf2f8; border-left: 5px solid #3498db; padding: 15px; margin: 15px 0; border-radius: 4px; font-size: 14px; }
    .step-intro b { color: #2c3e50; }
    .formula { background: #fff9e6; padding: 8px 12px; border-radius: 4px; font-family: 'Consolas', monospace; font-size: 13px; display: inline-block; margin: 5px 0; }
    @media print { .no-print { display: none; } }
    </style>
    """

    html = f"""<!DOCTYPE html><html lang="ru"><head><meta charset="UTF-8"><title>Статистический отчёт</title>{css}</head><body>
    <h1> Статистический отчёт: анализ «До — После»</h1>
    <div class="meta">
        <p><b>Файл данных:</b> {results_dict.get('file_name', '—')}</p>
        <p><b>Дата анализа:</b> {date_str}</p>
        <p><b>Всего спортсменов:</b> {results_dict.get('n_total', '—')}</p>
        <p><b>Контрольная группа</b> («{results_dict.get('group_ctrl', '—')}»): n = {results_dict.get('n_ctrl', '—')}</p>
        <p><b>Экспериментальная группа</b> («{results_dict.get('group_exp', '—')}»): n = {results_dict.get('n_exp', '—')}</p>
    </div>
    <button class="no-print" onclick="window.print()" style="padding:10px 20px; background:#3498db; color:white; border:none; border-radius:4px; cursor:pointer; font-size:14px; margin-bottom: 20px;">
         Распечатать / Сохранить в PDF
    </button>

    <div class="step-intro">
        <b> Как читать отчёт:</b><br>
        Каждый этап отвечает на свой вопрос исследования
        и логически подводит к следующему. В начале каждого раздела дано <b>пояснение простыми словами</b> —
        что мы проверяем, зачем и как интерпретировать результат. В конце отчёта собраны <b>итоговые выводы</b>,
        отвечающие на главный вопрос: <i>«Есть ли эффект от применённой методики?»</i>
    </div>"""

    # ═══════════════════════════════════════════════════════════════
    # ЭТАП 1. ОПИСАТЕЛЬНАЯ СТАТИСТИКА
    # ═══════════════════════════════════════════════════════════════
    html += """
    <h2>Этап 1. Описательная статистика</h2>
    <div class="step-intro">
        <b> Цель этапа:</b> получить первичное представление о данных — каковы типичные значения,
        насколько они разбросаны, есть ли явные выбросы. Это «портрет» выборки до любых статистических тестов.
    </div>
    <div class="explanation">
        <b>Что показывает таблица:</b><br>
        • <b>n</b> — объём выборки (сколько значений использовано в расчётах).<br>
        • <b>M (среднее арифметическое)</b> — «центр» данных. Чувствительно к выбросам.<br>
        • <b>SD (стандартное отклонение)</b> — насколько значения отклоняются от среднего.
        Большое SD = данные сильно разбросаны.<br>
        • <b>Md (медиана)</b> — значение, делящее выборку пополам. Устойчива к выбросам,
        поэтому при ненормальном распределении именно она характеризует «типичного» спортсмена.<br>
        • <b>Q1 и Q3 (квартили)</b> — границы средних 50% данных. <b>IQR = Q3 − Q1</b> — межквартильный размах.<br>
        • <b>SE (стандартная ошибка среднего)</b> — показывает, насколько точно среднее выборки
        оценивает среднее генеральной совокупности. Чем больше n, тем меньше SE.
    </div>
    <div class="explanation">
        <b>На что смотреть:</b> Сравните <b>M и Md</b>. Если они близки — распределение симметрично.
        Если сильно расходятся — есть асимметрия или выбросы. Сравните <b>SD и IQR</b> между группами:
        если в одной группе разброс сильно больше — это важно для выбора теста.
    </div>"""
    if 'desc_table' in results_dict:
        html += df_to_html(results_dict['desc_table'])

    # ═══════════════════════════════════════════════════════════════
    # ЭТАП 2. ПРОВЕРКА НОРМАЛЬНОСТИ
    # ═══════════════════════════════════════════════════════════════
    html += """
    <h2>Этап 2. Проверка нормальности распределения</h2>
    <div class="step-intro">
        <b> Цель этапа:</b> определить, какой математический аппарат применим к нашим данным.
        Статистические тесты делятся на <b>параметрические</b> (требуют нормального распределения)
        и <b>непараметрические</b> (работают с любыми распределениями). Этот этап — «переключатель» между ними.
    </div>
    <div class="explanation">
        <b>Что такое нормальное распределение:</b> Это симметричное «колоколообразное» распределение,
        в котором большинство значений сосредоточено вокруг среднего, а крайние значения встречаются редко.
        Многие статистические тесты (t-тест, ANOVA) <i>выведены в предположении</i>, что данные нормальны.
    </div>
    <div class="explanation">
        <b>Критерий Шапиро-Уилка (Shapiro-Wilk):</b><br>
        • <b>H₀ (нулевая гипотеза):</b> данные распределены нормально.<br>
        • <b>H₁ (альтернативная):</b> распределение отличается от нормального.<br>
        • <b>p ≥ 0.05</b> → нет оснований отвергать H₀ → распределение <b>можно считать нормальным</b>.<br>
        • <b>p &lt; 0.05</b> → отвергаем H₀ → распределение <b>не нормальное</b>.<br>
        <b>Критерий Дэвида-Шертона (D'Agostino-Pearson)</b> — дополнительная проверка, более чувствительная
        к асимметрии и эксцессу.
    </div>
    <div class="explanation">
        <b>Почему это важно:</b> Если применить параметрический тест к ненормальным данным,
        p-value может быть некорректным — мы рискуем сделать ложный вывод. Поэтому программа
        <b>автоматически</b> выбирает подходящий тест на основе результатов этого этапа.
    </div>"""
    if 'norm_table' in results_dict:
        html += df_to_html(results_dict['norm_table'], pvalue_cols=['Shapiro p', 'DP p'])
    if results_dict.get('all_normal'):
        html += """<div class="conclusion">✅ <b>Вывод этапа:</b> все распределения можно считать нормальными.
        Применены параметрические критерии (t-тесты) — они обладают большей статистической мощностью
        (способны обнаружить даже небольшие эффекты).</div>"""
    else:
        html += """<div class="warning">⚠ <b>Вывод этапа:</b> не все распределения нормальны.
        Применены непараметрические критерии (Wilcoxon, Mann-Whitney) — они не требуют нормальности,
        но работают с <i>рангами</i> (порядком значений), а не с их абсолютными величинами.</div>"""

    # ═══════════════════════════════════════════════════════════════
    # ЭТАП 3. ВНУТРИГРУППОВЫЕ СРАВНЕНИЯ
    # ═══════════════════════════════════════════════════════════════
    html += """
    <h2>Этап 3. Внутригрупповые сравнения (До vs После)</h2>
    <div class="step-intro">
        <b> Цель этапа:</b> ответить на вопрос — <i>«изменились ли показатели внутри каждой группы
        после воздействия?»</i> Мы сравниваем каждого спортсмена «самого с собой»: его результат
        до и после. Это <b>парные (зависимые) выборки</b>.
    </div>
    <div class="explanation">
        <b>Какие тесты используются:</b><br>
        • <b>Парный t-тест (Стьюдента)</b> — параметрический. Проверяет, отличается ли среднее
        разностей (После − До) от нуля. Применим, когда разности распределены нормально.<br>
        • <b>Критерий Уилкоксона (Wilcoxon signed-rank)</b> — непараметрический аналог.
        Работает с рангами разностей, устойчив к выбросам.
    </div>
    <div class="explanation">
        <b>Как читать p-value:</b><br>
        • <b>p &lt; 0.05</b> → изменения статистически значимы (маловероятно, что они случайны).<br>
        • <b>p ≥ 0.05</b> → изменения недостоверны (могли произойти случайно).<br>
        <b>Звёздочки:</b> *** — p&lt;0.001 (высоко значимо), ** — p&lt;0.01, * — p&lt;0.05.
    </div>
    <div class="explanation">
        <b>Размер эффекта — почему одного p-value недостаточно:</b><br>
        p-value зависит от размера выборки: при большом n даже крошечные изменения становятся «значимыми».
        <b>Размер эффекта</b> измеряет <i>практическую</i> значимость — насколько сильно изменились показатели,
        независимо от числа спортсменов.<br>
        • <b>d Коэна</b> (для t-теста): 0.2 — малый, 0.5 — средний, 0.8 — большой эффект.<br>
        • <b>r Розенталя</b> (для Уилкоксона): 0.1 — малый, 0.3 — средний, 0.5 — большой.
    </div>"""
    if 'within_table' in results_dict:
        html += df_to_html(results_dict['within_table'], pvalue_cols=['p-value'])

    # ═══════════════════════════════════════════════════════════════
    # ЭТАП 4. МЕЖГРУППОВЫЕ СРАВНЕНИЯ
    # ═══════════════════════════════════════════════════════════════
    html += """
    <h2>Этап 4. Межгрупповые сравнения</h2>
    <div class="step-intro">
        <b> Цель этапа:</b> ответить на два ключевых вопроса:<br>
        1) <i>«Были ли группы одинаковы в начале эксперимента?»</i> (сравнение «До») — это проверка
        корректности формирования групп.<br>
        2) <i>«Отличается ли эффект методики между группами?»</i> (сравнение изменений) — это
        <b>главный вопрос исследования</b>.
    </div>
    <div class="explanation">
        <b>Строка «До (между группами)»:</b> Если p ≥ 0.05 — группы были сопоставимы в начале,
        и любые последующие различия можно связывать с воздействием. Если p &lt; 0.05 —
        группы изначально различались, и это нужно учитывать (например, через ANCOVA).
    </div>
    <div class="explanation">
        <b>Строка «Изменения (между группами)» — ГЛАВНАЯ:</b> Показывает, различается ли
        <i>величина сдвига</i> между контрольной и экспериментальной группами.<br>
        • <b>p &lt; 0.05</b> → методика работает по-разному в группах → эффект доказан.<br>
        • <b>p ≥ 0.05</b> → различия между группами недостоверны → эффект методики не подтверждён.
    </div>
    <div class="explanation">
        <b>Размер эффекта для независимых выборок:</b><br>
        • <b>d Коэна</b> (параметрический): 0.2 / 0.5 / 0.8 — малый / средний / большой.<br>
        • <b>r Розенталя</b> — аналог для Mann-Whitney.<br>
        • <b>Δ Клиффа (Cliff's delta)</b> — вероятность того, что случайно выбранный спортсмен
        из экспериментальной группы покажет лучший результат, чем случайно выбранный из контрольной,
        минус обратная вероятность. Диапазон от −1 до +1. 0.147 / 0.33 / 0.474 — малый / средний / большой.
    </div>"""
    if 'between_table' in results_dict:
        html += df_to_html(results_dict['between_table'], pvalue_cols=['p-value'])

    # ═══════════════════════════════════════════════════════════════
    # ЭТАП 5. ANOVA / ANCOVA
    # ═══════════════════════════════════════════════════════════════
    if results_dict.get('anova_table') is not None:
        html += """
        <h2>Этап 5. Математическая модель (ANOVA / ANCOVA)</h2>
        <div class="step-intro">
            <b> Цель этапа:</b> построить единую математическую модель, которая одновременно учитывает
            влияние <b>группы</b>, <b>времени</b> и их <b>взаимодействия</b>. Это наиболее строгий
            и информативный способ анализа.
        </div>
        <div class="explanation">
            <b>Двухфакторная ANOVA (дисперсионный анализ):</b><br>
            Модель: <span class="formula">Результат ~ Группа + Время + Группа×Время</span><br>
            • <b>Эффект «Группа»</b> — различаются ли группы в принципе (независимо от времени).<br>
            • <b>Эффект «Время»</b> — меняются ли показатели у всех спортсменов (До vs После).<br>
            • <b>Взаимодействие «Группа×Время»</b> — <b>ключевой показатель</b>. Если оно значимо (p&lt;0.05),
            это означает, что динамика показателей <i>зависит от принадлежности к группе</i>.
            Это и есть <b>«золотой стандарт»</b> доказательства эффективности методики.
        </div>
        <div class="explanation">
            <b>ANCOVA (ковариационный анализ):</b><br>
            Модель: <span class="formula">После ~ Группа + До</span><br>
            Более строгий анализ, который «очищает» результат «После» от влияния исходного уровня «До».
            Полезен, если группы изначально немного различались. Коэффициент при переменной «Группа»
            показывает <b>чистый эффект методики</b>, очищенный от исходных различий.
        </div>
        <h3>5a. Двухфакторная ANOVA</h3>"""
        html += df_to_html(results_dict['anova_table'], pvalue_cols=['p-value'])
        if results_dict.get('ancova_table') is not None:
            html += "<h3>5b. ANCOVA (контроль исходного уровня)</h3>"
            html += df_to_html(results_dict['ancova_table'], pvalue_cols=['P>|t|'])

    # ═══════════════════════════════════════════════════════════════
    # ЭТАП 6. ГРАФИКИ
    # ═══════════════════════════════════════════════════════════════
    html += """
    <h2>Этап 6. Визуализация результатов</h2>
    <div class="step-intro">
        <b> Цель этапа:</b> наглядно подтвердить числовые выводы. Графики позволяют увидеть то,
        что не всегда очевидно из таблиц: форму распределения, индивидуальные траектории, наличие
        выбросов, характер связи между «До» и «После».
    </div>
    <div class="explanation">
        <b>Как читать графики:</b> Каждый рисунок ниже сопровождается подписью, объясняющей,
        <i>что именно</i> он показывает и <i>на что обращать внимание</i>. Графики расположены
        в порядке от общих (распределения по группам) к частным (индивидуальные траектории,
        связь между измерениями).
    </div>"""

    captions = {
        '01_violin_before_after': (
            'Скрипичная диаграмма (Violin plot). Показывает <b>форму распределения</b> данных '
            'До и После в каждой группе. Ширина «скрипки» — плотность значений: где шире, '
            'там больше спортсменов. Обратите внимание: сместился ли «центр» скрипки от «До» к «После» '
            'и изменилась ли её форма (разброс). Сравните направление сдвига между контрольной '
            'и экспериментальной группами.'
        ),
        '02_boxplot_changes': (
            'Ящик с усами (Boxplot) для <b>индивидуальных изменений</b> (После − До). '
            'Линия внутри ящика — <b>медиана изменений</b>. Границы ящика — Q1 и Q3 (средние 50% спортсменов). '
            '«Усы» — крайние значения. Красная пунктирная линия — ноль (отсутствие изменений). '
            '<b>Если ящик расположен выше нуля — группа в среднем улучшилась; ниже — ухудшилась.</b> '
            'Точки за усами — возможные выбросы.'
        ),
        '03_spaghetti_plot': (
            'График «Спагетти» — <b>индивидуальные траектории</b> каждого спортсмена от «До» к «После». '
            'Каждая линия — один человек. Толстые линии — средние значения по группам. '
            'Этот график показывает <b>вариативность эффекта</b>: у всех ли спортсменов динамика '
            'одинакова, или есть те, кто reagирует сильно, а кто-то не reagирует вовсе. '
            'Если линии в экспериментальной группе идут преимущественно вверх, а в контрольной — '
            'горизонтально, это визуальное подтверждение эффекта методики.'
        ),
        '04_barplot_medians_iqr': (
            'Столбчатая диаграмма <b>медиан</b> с «усами», показывающими межквартильный размах (IQR). '
            'В отличие от средних, медиана устойчива к выбросам, поэтому этот график надёжно '
            'отражает «типичного» спортсмена. Если столбцы «После» заметно выше «До» только '
            'в экспериментальной группе — это визуальное подтверждение эффекта.'
        ),
        '05_qqplot_control': (
            'QQ-plot для <b>контрольной группы</b>. Проверяет нормальность распределения изменений. '
            'Если точки лежат <b>близко к прямой линии</b> — распределение близко к нормальному. '
            'Отклонения на концах («тяжёлые хвосты») говорят о выбросах. S-образная кривая — '
            'признак асимметрии.'
        ),
        '06_qqplot_experimental': (
            'QQ-plot для <b>экспериментальной группы</b>. Интерпретация аналогична предыдущему. '
            'Сравните оба QQ-plot: если в одной группе точки сильнее отклоняются от прямой — '
            'там распределение менее нормальное, и непараметрические тесты более оправданы.'
        ),
        '07_histogram_control': (
            'Гистограмма распределения изменений в <b>контрольной группе</b>. По оси X — величина '
            'изменения, по оси Y — сколько спортсменов попали в этот диапазон. Плавная кривая (KDE) — '
            'оценка плотности распределения. Красная линия — медиана изменений. '
            'Симметричная форма вокруг нуля говорит об отсутствии систематического сдвига.'
        ),
        '08_histogram_experimental': (
            'Гистограмма распределения изменений в <b>экспериментальной группе</b>. '
            '<b>Сравните её с гистограммой контрольной группы:</b> если центр экспериментальной '
            'гистограмры смещён вправо (в сторону положительных значений) — методика работает. '
            'Красная линия — медиана изменений.'
        ),
        '09_scatter_before_after': (
            'Диаграмма рассеяния «До» vs «После». Каждая точка — один спортсмен. '
            'Пунктирная линия <b>y = x</b> — «линия отсутствия изменений»: точки на ней означают, '
            'что результат не изменился. <b>Точки выше линии — улучшение, ниже — ухудшение.</b> '
            'Наклонные линии — тренды для каждой группы. Если тренд экспериментальной группы '
            'круче и расположен выше линии y=x — методика эффективна.'
        ),
        '10_summary_effects': (
            'Сводный точечный график <b>средних значений с 95% доверительными интервалами</b>. '
            'Точки — средние, «усы» — 95% ДИ (диапазон, в котором с вероятностью 95% находится '
            'истинное среднее генеральной совокупности). <b>Если доверительные интервалы двух групп '
            'не перекрываются — различия, как правило, статистически значимы.</b> '
            'Этот график — финальная визуальная сводка всего исследования.'
        )
    }
    for key, caption in captions.items():
        if key in results_dict.get('figures', {}):
            html += f"""
            <div class="figure">
                <img src="{results_dict['figures'][key]}" alt="{caption}">
                <div class="figure-caption"><b>Рис.</b> {caption}</div>
            </div>"""

    # ═══════════════════════════════════════════════════════════════
    # ЭТАП 7. ИТОГОВЫЕ ВЫВОДЫ
    # ═══════════════════════════════════════════════════════════════
    html += """
    <h2>Этап 7. Итоговые выводы</h2>
    <div class="step-intro">
        <b> Цель этапа:</b> собрать воедино все числовые и визуальные доказательства и дать
        окончательный ответ на главный вопрос исследования.
    </div>"""
    for conclusion in results_dict.get('conclusions', []):
        html += f'<div class="conclusion">{conclusion}</div>'

    html += """
    <div class="explanation" style="margin-top: 30px;">
        <b> Общая интерпретация результатов:</b><br>
        Для того чтобы утверждать, что <b>методика эффективна</b>, необходимо одновременное выполнение
        нескольких условий:<br>
        1) Группы были сопоставимы в начале (p ≥ 0.05 в сравнении «До»).<br>
        2) В экспериментальной группе есть достоверные сдвиги (p&lt;0.05 во внутригрупповом тесте).<br>
        3) Эти сдвиги значимо больше, чем в контрольной группе (p&lt;0.05 в сравнении изменений).<br>
        4) Взаимодействие «Группа×Время» в ANOVA значимо (p&lt;0.05).<br>
        5) Размер эффекта — не менее средний (d ≥ 0.5 или r ≥ 0.3) — показывает <i>практическую</i> значимость.<br>
        Если все эти условия выполнены — эффект методики считается <b>доказанным</b>.
    </div>
    <hr style="margin-top: 40px;">
    <p style="text-align:center; color:#95a5a6; font-size:12px;">
        Отчёт сгенерирован автоматически программой статистического анализа.
        Рекомендуется проконсультироваться со специалистом для интерпретации результатов в контексте предметной области.
    </p></body></html>"""

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    return output_path