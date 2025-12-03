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
# app.py
import io
import streamlit as st
import plotly.express as px
import pandas as pd

from data_loader import load_excels, get_row_bounds, slice_range
from classification import classify_questions, QuestionType
from summary import build_all_summaries
from excel_export import build_excel_report
from pdf_export import build_pdf_report
from docx_export import build_docx_report
from pptx_export import build_pptx_report

st.set_page_config(
    page_title="Обробка результатів студентських опитувань",
    layout="wide",
)

# ----------------- допоміжні функції -----------------

def init_state():
    defaults = {
        "uploaded_files_store": None,  # [{"name":..., "bytes":...}, ...]
        "ld": None,
        "sliced": None,
        "qinfo": None,
        "summaries": None,
        "processed": False,
        "selected_code": None,
        "from_row": None,
        "to_row": None,
        "uploaded_names_snapshot": None,  # для порівняння, чи файли змінились
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

def store_uploaded_files(uploaded_files):
    """
    Зчитати байти з UploadedFile і зберегти в session_state.
    """
    if not uploaded_files:
        return
    store = []
    names = []
    for f in uploaded_files:
        try:
            b = f.read()
            name = getattr(f, "name", None) or "file"
            store.append({"name": name, "bytes": b})
            names.append(name)
        except Exception as e:
            st.error(f"Не вдалося зчитати файл {getattr(f,'name','')}: {e}")
            return
    st.session_state.uploaded_files_store = store
    st.session_state.uploaded_names_snapshot = names

def get_uploaded_files_from_store():
    """
    Повернути список file-like об'єктів (BytesIO) для load_excels.
    """
    if not st.session_state.uploaded_files_store:
        return None
    objs = []
    for item in st.session_state.uploaded_files_store:
        bio = io.BytesIO(item["bytes"])
        bio.name = item["name"]
        objs.append(bio)
    return objs

def clear_processing_state(keep_files=True):
    """
    Очищає оброблені результати; якщо keep_files=False, то також очищає збережені файли.
    """
    st.session_state.ld = None
    st.session_state.sliced = None
    st.session_state.qinfo = None
    st.session_state.summaries = None
    st.session_state.processed = False
    st.session_state.selected_code = None
    st.session_state.from_row = None
    st.session_state.to_row = None
    if not keep_files:
        st.session_state.uploaded_files_store = None
        st.session_state.uploaded_names_snapshot = None

def process_range_from_store(from_row: int, to_row: int):
    """
    Виконує обробку на базі байтів з uploaded_files_store.
    """
    files_for_load = get_uploaded_files_from_store()
    if not files_for_load:
        raise ValueError("Немає збережених файлів для обробки.")
    ld = load_excels(files_for_load)
    sliced = slice_range(ld, int(from_row), int(to_row))
    qinfo = classify_questions(sliced, technical_columns=1)
    summaries = build_all_summaries(sliced, qinfo)

    st.session_state.ld = ld
    st.session_state.sliced = sliced
    st.session_state.qinfo = qinfo
    st.session_state.summaries = summaries
    st.session_state.processed = True
    st.session_state.from_row = int(from_row)
    st.session_state.to_row = int(to_row)

    codes = [qs.question.code for qs in summaries]
    st.session_state.selected_code = codes[0] if codes else None

# ----------------- UI -----------------

def main():
    init_state()

    st.title("Система обробки результатів студентських опитувань")
    # повертаємо допоміжний опис (той самий, що був раніше)
    st.markdown(
        """
        Цей веб-застосунок призначений для аналізу результатів опитувань,
        проведених за допомогою Google Forms. Завантажте один або кілька
        файлів Excel, оберіть діапазон рядків і натисніть «Обробити діапазон».
        Після обробки ви зможете вибирати питання для перегляду таблиці
        та діаграми без повторної обробки.
        """
    )

    st.sidebar.header("Налаштування аналізу")

    uploaded_files = st.sidebar.file_uploader(
        "Завантажте файл(и) Excel з відповідями",
        type=["xlsx"],
        accept_multiple_files=True,
        help="Підтримується кілька файлів Google Forms; вони будуть об'єднані."
    )

    # Кнопки: обробити та скидання
    process_button = st.sidebar.button("Обробити діапазон")
    reset_button = st.sidebar.button("Скинути обробку")

    # Обробка нового завантаження: якщо користувач обрав файли у file_uploader,
    # імена відрізняються від snapshot -> зберегти байти і скинути попередню обробку
    if uploaded_files:
        names_now = [getattr(f, "name", "") for f in uploaded_files]
        if st.session_state.uploaded_names_snapshot != names_now:
            store_uploaded_files(uploaded_files)
            # після зміни файлів потрібно очистити оброблені результати
            clear_processing_state(keep_files=True)

    # Якщо немає збережених файлів у сесії — підказка і вихід
    if not st.session_state.uploaded_files_store:
        st.info("Завантажте хоча б один файл Excel у боковій панелі, щоб розпочати аналіз.")
        return

    # Показати короткий огляд (якщо вже оброблено раніше)
    if st.session_state.processed and st.session_state.ld:
        ld = st.session_state.ld
        st.subheader("Короткий огляд останньо оброблених відповідей")
        st.write(f"Кількість рядків: **{ld.n_rows}**, стовпців: **{ld.n_cols}**")
        st.dataframe(ld.df.head())

        # для зручності даємо можливість змінити діапазон (але обробка відбудеться лише по натисканню)
        min_row, max_row = get_row_bounds(ld)
        from_row = st.sidebar.number_input(
            "Від (рядок)", min_value=min_row, max_value=max_row,
            value=st.session_state.from_row or min_row, step=1
        )
        to_row = st.sidebar.number_input(
            "До (рядок)", min_value=min_row, max_value=max_row,
            value=st.session_state.to_row or max_row, step=1
        )

        if process_button:
            try:
                process_range_from_store(from_row, to_row)
                st.success("Обробку завершено.")
            except Exception as e:
                st.error(f"Помилка обробки: {e}")
                return

    else:
        # якщо ще не обробляли — тимчасово прочитати файли, щоб показати допустимі межі
        try:
            tmp_objs = get_uploaded_files_from_store()
            tmp_ld = load_excels(tmp_objs)
            min_row, max_row = get_row_bounds(tmp_ld)
        except Exception as e:
            st.error(f"Помилка читання файлів: {e}")
            return

        from_row = st.sidebar.number_input("Від (рядок)", min_value=min_row, max_value=max_row, value=min_row, step=1)
        to_row = st.sidebar.number_input("До (рядок)", min_value=min_row, max_value=max_row, value=max_row, step=1)

        if process_button:
            try:
                process_range_from_store(from_row, to_row)
                st.success("Обробку завершено.")
            except Exception as e:
                st.error(f"Помилка обробки: {e}")
                return
        else:
            st.info("Оберіть діапазон та натисніть «Обробити діапазон».")

    # Кнопка reset очищує лише обробку (або можна змінити, щоб очищати і файли)
    if reset_button:
        clear_processing_state(keep_files=False)
        st.success("Обробку скинуто. Можна обрати новий діапазон або завантажити інші файли.")
        st.rerun()

    # --- Після успішної обробки показуємо результати ---
    if st.session_state.processed:
        summaries = st.session_state.summaries or []
        qinfo = st.session_state.qinfo or {}
        sliced = st.session_state.sliced

        st.subheader("Класифікація запитань")
        q_table = pd.DataFrame(
            [
                {"Код": info.code, "Назва стовпця": col, "Тип": info.qtype.value}
                for col, info in qinfo.items()
                if info.qtype != QuestionType.TECHNICAL
            ]
        )
        st.dataframe(q_table)

        st.subheader("Таблиці відповідей за всіма запитаннями")
        for qs in summaries:
            with st.expander(f"{qs.question.code}. {qs.question.text}"):
                st.dataframe(qs.table)

        # вибір питання для діаграми — використовуємо session_state.selected_code
        st.subheader("Окреме питання (діаграма)")

        available_codes = [qs.question.code for qs in summaries]
        if not available_codes:
            st.warning("Немає запитань з варіантами для побудови діаграм.")
        else:
            # встановлюємо індекс для selectbox згідно з останнім обраним кодом
            default_index = 0
            if st.session_state.selected_code and st.session_state.selected_code in available_codes:
                default_index = available_codes.index(st.session_state.selected_code)

            selected_code = st.selectbox(
                "Оберіть код питання:",
                options=available_codes,
                index=default_index,
                key="select_question_box"
            )

            # зберегти вибір у сесії, але не тригерити повторну обробку
            st.session_state.selected_code = selected_code

            selected = next((qs for qs in summaries if qs.question.code == st.session_state.selected_code), None)

            if selected is None or selected.table.empty:
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

        # --- Експорт звіту ---
        st.subheader("Експорт результатів")
        range_info = f"Рядки {st.session_state.from_row}–{st.session_state.to_row} (усього {len(sliced)} анкет)"
        
        # Створюємо 3 колонки для кнопок
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            # EXCEL
            report_bytes_xlsx = build_excel_report(
                original_df=st.session_state.ld.df,
                sliced_df=st.session_state.sliced,
                qinfo=st.session_state.qinfo,
                summaries=st.session_state.summaries,
                range_info=range_info,
            )
            st.download_button(
                label="📥 Excel звіт",
                data=report_bytes_xlsx,
                file_name="survey_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        with col2:
            # PDF
            if st.button("📄 PDF звіт"):
                with st.spinner("Генеруємо PDF..."):
                    try:
                        from pdf_export import build_pdf_report # lazy import
                        report_bytes_pdf = build_pdf_report(
                            original_df=st.session_state.ld.df,
                            sliced_df=st.session_state.sliced,
                            summaries=st.session_state.summaries,
                            range_info=range_info,
                        )
                        st.download_button(
                            label="📥 Завантажити PDF",
                            data=report_bytes_pdf,
                            file_name="survey_results.pdf",
                            mime="application/pdf",
                        )
                    except Exception as e:
                        st.error(f"Error: {e}")

        with col3:
            # WORD (DOCX)
            if st.button("📝 Word звіт"):
                with st.spinner("Генеруємо DOCX..."):
                    try:
                        report_bytes_docx = build_docx_report(
                            original_df=st.session_state.ld.df,
                            sliced_df=st.session_state.sliced,
                            summaries=st.session_state.summaries,
                            range_info=range_info,
                        )
                        st.download_button(
                            label="📥 Завантажити Word",
                            data=report_bytes_docx,
                            file_name="survey_results.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        )
                    except Exception as e:
                         st.error(f"Error DOCX: {e}")
                
        with col4:
            # PPTX
            if st.button("🖥️ PPTX звіт"):
                with st.spinner("Генеруємо PowerPoint..."):
                    try:
                        report_bytes_pptx = build_pptx_report(
                            original_df=st.session_state.ld.df,
                            sliced_df=st.session_state.sliced,
                            summaries=st.session_state.summaries,
                            range_info=range_info,
                        )
                        st.download_button(
                            label="📥 Завантажити PPTX",
                            data=report_bytes_pptx,
                            file_name="survey_results.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        )
                    except Exception as e:
                        st.error(f"Error PPTX: {e}")

if __name__ == "__main__":
    main()
