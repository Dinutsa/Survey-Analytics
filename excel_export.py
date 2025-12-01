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

    :param original_df: повна таблиця відповідей.
    :param sliced_df: вибраний користувачем діапазон.
    :param qinfo: інформація про питання.
    :param summaries: список зведених таблиць.
    :param range_info: текстовий опис діапазону (для титулу).
    """
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # 1. Технічна інформація
        meta_df = pd.DataFrame(
            {
                "Параметр": [
                    "Загальна кількість відповідей",
                    "Кількість відповідей у вибраному діапазоні",
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

        # 2. Вихідні дані (скорочено)
        sliced_df.to_excel(writer, sheet_name="Вихідні_дані", index=False)

        # 3. Таблиці підсумків (по всіх питаннях)
        workbook = writer.book
        ws = workbook.add_worksheet("Підсумки")
        writer.sheets["Підсумки"] = ws

        row = 0
        for qs in summaries:
            ws.write(row, 0, f"{qs.question.code}. {qs.question.text}")
            row += 1
            # записуємо таблицю
            qs.table.to_excel(
                writer,
                sheet_name="Підсумки",
                startrow=row,
                startcol=0,
                index=False,
                header=True,
            )
            row += len(qs.table) + 2  # відступ між блоками
        chart = workbook.add_chart({'type': 'pie'})
        chart.add_series({
            'name':       [sheet_name, 0, 0],  # Назва діаграми
            'categories': [sheet_name, 1, 0, 5, 0],  # Діапазон підписів
            'values':     [sheet_name, 1, 1, 5, 1],  # Діапазон значень
            })
        ws.insert_chart('E2', chart)

    output.seek(0)
    return output.read()
