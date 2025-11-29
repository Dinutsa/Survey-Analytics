"""
Головний модуль веб-застосунку для обробки результатів студентських опитувань.

Забезпечує:
- завантаження файлів Excel із Google Forms;
- вибір діапазону рядків;
- класифікацію типів питань;
- обчислення частот і відсотків;
- побудову кругових діаграм;
- формування Excel-звіту.

Такий підхід відповідає рекомендаціям з автоматизації обробки
соціологічних досліджень та потребам систем внутрішнього забезпечення
якості освіти ЗВО.
"""
"працюйй"
from __future__ import annotations

import streamlit as st
import plotly.express as px
import pandas as pd

from data_loader import load_excels, get_row_bounds, slice_range
from classification import classify_questions, QuestionType
from summary import build_all_summaries
from excel_export import build_excel_report


st.set_page_config(
    page_title="Обробка результатів студентських опитувань",
    layout="wide",
)


def main() -> None:
    st.title("Система обробки результатів студентських опитувань")

    st.markdown(
        """
        Цей веб-застосунок призначений для аналізу результатів опитувань,
        проведених за допомогою Google Forms. Завантажте один або кілька
        файлів Excel, оберіть діапазон анкет, а система автоматично
        сформує таблиці, діаграми та звіт у форматі Excel.
        """
    )

    # --- Бокова панель ---
    st.sidebar.header("Налаштування аналізу")

    uploaded_files = st.sidebar.file_uploader(
        "Завантажте файл(и) Excel з відповідями",
        type=["xlsx"],
        accept_multiple_files=True,
    )

    process_button = st.sidebar.button("Обробити діапазон")

    if not uploaded_files:
        st.info("Завантажте хоча б один файл Excel, щоб розпочати аналіз.")
        return

    # Завантаження та попередній перегляд
    try:
        ld = load_excels(uploaded_files)
    except ValueError as e:
        st.error(str(e))
        return

    st.subheader("Короткий огляд об’єднаних відповідей")
    st.write(f"Кількість рядків: **{ld.n_rows}**, стовпців: **{ld.n_cols}**")
    st.dataframe(ld.df.head())

    min_row, max_row = get_row_bounds(ld)

    st.sidebar.markdown("### Діапазон анкет для обробки")
    from_row = st.sidebar.number_input(
        "Від (рядок)", min_value=min_row, max_value=max_row, value=min_row, step=1
    )
    to_row = st.sidebar.number_input(
        "До (рядок)", min_value=min_row, max_value=max_row, value=max_row, step=1
    )

    if not process_button:
        st.info("Оберіть діапазон і натисніть «Обробити діапазон».")
        return

    # --- Обробка діапазону ---
    try:
        sliced = slice_range(ld, int(from_row), int(to_row))
    except ValueError as e:
        st.error(str(e))
        return

    st.success(
        f"Діапазон оброблено. Рядки: {from_row}–{to_row}, кількість анкет: {len(sliced)}."
    )

    # Класифікація питань (пропускаємо перший стовпець – позначка часу)
    qinfo = classify_questions(sliced, technical_columns=1)

    # Таблиця з описом питань
    st.subheader("Класифікація запитань")
    q_table = pd.DataFrame(
        [
            {
                "Код": info.code,
                "Назва стовпця": col,
                "Тип": info.qtype.value,
            }
            for col, info in qinfo.items()
            if info.qtype != QuestionType.TECHNICAL
        ]
    )
    st.dataframe(q_table)

    # Розрахунок підсумків
    summaries = build_all_summaries(sliced, qinfo)

    st.subheader("Таблиці відповідей за всіма запитаннями")
    for qs in summaries:
        with st.expander(f"{qs.question.code}. {qs.question.text}"):
            st.dataframe(qs.table)

    # --- Окрема діаграма по вибраному питанню ---
    st.subheader("Окреме питання (діаграма)")

    available_codes = [qs.question.code for qs in summaries]
    if not available_codes:
        st.warning("Немає запитань з варіантами для побудови діаграм.")
    else:
        selected_code = st.selectbox("Оберіть код питання:", available_codes)
        selected = next(qs for qs in summaries if qs.question.code == selected_code)

        if selected.table.empty:
            st.warning("Для цього питання немає даних для побудови діаграми.")
        else:
            fig = px.pie(
                selected.table,
                names="Варіант відповіді",
                values="Кількість",
                title=f"{selected.question.code}. {selected.question.text}",
                hole=0.0,
            )
            st.plotly_chart(fig, use_container_width=True)
            st.dataframe(selected.table)

    # --- Завантаження Excel-звіту ---
    st.subheader("Експорт результатів")

    range_info = f"Рядки {from_row}–{to_row} (усього {len(sliced)} анкет)"
    report_bytes = build_excel_report(
        original_df=ld.df,
        sliced_df=sliced,
        qinfo=qinfo,
        summaries=summaries,
        range_info=range_info,
    )

    st.download_button(
        label="Завантажити звіт (Excel)",
        data=report_bytes,
        file_name="oputuvalny_analiz_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
