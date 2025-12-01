"""
Модуль формування Excel-звіту з результатами опитування.

Застосування звітів у форматі XLSX відповідає практиці документування
результатів соціологічних досліджень у ЗВО та вимогам внутрішніх
систем забезпечення якості освіти.
"""
"""
Модуль формування Excel-звіту з результатами опитування.

Застосування звітів у форматі XLSX відповідає практиці документування
результатів соціологічних досліджень у ЗВО та вимогам внутрішніх
систем забезпечення якості освіти.
"""

from __future__ import annotations

import io
from typing import Dict, List

import pandas as pd

from classification import QuestionInfo
from summary import QuestionSummary


def build_excel_report(
    original_df: pd.DataFrame,
    sliced_df: pd.DataFrame,
    qinfo: Dict[str, QuestionInfo],
    summaries: List[QuestionSummary],
    range_info: str,
) -> bytes:
    """
    Створює Excel-звіт та повертає його у вигляді байтів для завантаження.
    Додає кругові діаграми для питань, що мають зведені таблиці.

    :param original_df: повна таблиця відповідей.
    :param sliced_df: вибраний користувачем діапазон.
    :param qinfo: інформація про питання.
    :param summaries: список зведених таблиць.
    :param range_info: текстовий опис діапазону (для титулу).
    """
    output = io.BytesIO()

    # Використовуємо engine="xlsxwriter", бо він підтримує створення діаграм
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # --- 1. Технічна інформація ---
        meta_df = pd.DataFrame(
            {
                "Параметр": [
                    "Загальна кількість анкет у файлі",
                    "Кількість анкет у вибраному діапазоні",
                    "Діапазон обробки",
                ],
                "Значення": [
                    len(original_df),
                    len(sliced_df),
                    range_info,
                ],
            }
        )
        meta_df.to_excel(writer, sheet_name="Технічна_інформація", index=False)

        # Підлаштуємо ширину колонок для краси
        workbook = writer.book
        worksheet_meta = writer.sheets["Технічна_інформація"]
        worksheet_meta.set_column("A:A", 40)
        worksheet_meta.set_column("B:B", 50)

        # --- 2. Вихідні дані (скорочено/фільтровано) ---
        sliced_df.to_excel(writer, sheet_name="Вихідні_дані", index=False)

        # --- 3. Таблиці підсумків та діаграми ---
        sheet_name_summary = "Підсумки"
        ws = workbook.add_worksheet(sheet_name_summary)
        writer.sheets[sheet_name_summary] = ws

        # Формат для заголовків питань
        bold_fmt = workbook.add_format({"bold": True, "font_size": 11})
        
        # Задаємо ширину колонок для підсумків
        ws.set_column("A:A", 40)  # Варіант відповіді
        ws.set_column("B:C", 10)  # Числа та %

        current_row = 0

        for qs in summaries:
            question_title = f"{qs.question.code}. {qs.question.text}"
            
            # 1. Записуємо назву питання
            ws.write(current_row, 0, question_title, bold_fmt)
            current_row += 1

            # Якщо таблиця порожня (наприклад, відкрите питання або немає даних)
            if qs.table.empty:
                ws.write(current_row, 0, "(Немає даних для діаграми або текстове питання)")
                current_row += 2
                continue

            # 2. Записуємо таблицю даних
            # qs.table має колонки: ["Варіант відповіді", "Кількість", "%"]
            # startrow=current_row, startcol=0
            qs.table.to_excel(
                writer,
                sheet_name=sheet_name_summary,
                startrow=current_row,
                startcol=0,
                index=False,
                header=True,
            )

            # Визначаємо координати даних для діаграми
            # Заголовок таблиці займає 1 рядок, тому дані починаються з current_row + 1
            n_rows = len(qs.table)
            first_data_row = current_row + 1
            last_data_row = current_row + n_rows

            # Колонки (0-based): A=0 (Labels), B=1 (Values)
            
            # 3. Створюємо діаграму
            chart = workbook.add_chart({"type": "pie"})
            
            chart.add_series({
                "name": "Кількість",
                # [sheetname, first_row, first_col, last_row, last_col]
                "categories": [sheet_name_summary, first_data_row, 0, last_data_row, 0],
                "values":     [sheet_name_summary, first_data_row, 1, last_data_row, 1],
                "data_labels": {"percentage": True},  # Показувати відсотки на діаграмі
            })
            
            chart.set_title({"name": question_title})
            chart.set_style(10)  # Стиль діаграми (можна змінювати 1-48)

            # 4. Вставляємо діаграму праворуч від таблиці (наприклад, колонка E, індекс 4)
            # Вставляємо трохи вище (на рівні заголовку питання), щоб було компактно
            ws.insert_chart(current_row - 1, 4, chart)

            # 5. Зсуваємо курсор вниз
            # Нам треба відступити місце або під таблицю, або під діаграму (що більше)
            # Стандартна висота діаграми ~15 рядків.
            rows_occupied = max(n_rows + 3, 18)
            current_row += rows_occupied

    return output.getvalue()