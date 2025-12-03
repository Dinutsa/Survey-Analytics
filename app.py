import io
import os
import streamlit as st
import plotly.express as px
import pandas as pd

# Імпорти логіки
from data_loader import load_excels, get_row_bounds, slice_range
from classification import classify_questions, QuestionType
from summary import build_all_summaries

# Імпорти експорту
from excel_export import build_excel_report
from pdf_export import build_pdf_report
from docx_export import build_docx_report
from pptx_export import build_pptx_report

st.set_page_config(
    page_title="Обробка результатів студентських опитувань",
    layout="wide",
)

def init_state():
    defaults = {
        "uploaded_files_store": None,
        "ld": None,
        "sliced": None,
        "qinfo": None,
        "summaries": None,
        "processed": False,
        "selected_code": None,
        "from_row": 0,
        "to_row": 0,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

st.title("Аналіз результатів опитувань (Google Forms)")

# --- SIDEBAR ---
with st.sidebar:
    st.header("1. Завантаження даних")
    uploaded_files = st.file_uploader(
        "Оберіть Excel-файли (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True
    )

    if uploaded_files:
        if st.session_state.ld is None or uploaded_files != st.session_state.uploaded_files_store:
            try:
                ld = load_excels(uploaded_files)
                st.session_state.ld = ld
                st.session_state.uploaded_files_store = uploaded_files
                
                min_r, max_r = get_row_bounds(ld)
                st.session_state.from_row = min_r
                st.session_state.to_row = max_r
                st.session_state.processed = False
            except Exception as e:
                st.error(f"Помилка: {e}")

    if st.session_state.ld:
        st.success(f"Завантажено: {st.session_state.ld.n_rows} анкет.")
        st.divider()
        st.header("2. Фільтрація")
        
        min_r, max_r = get_row_bounds(st.session_state.ld)
        if max_r > min_r:
            r_range = st.slider(
                "Діапазон рядків",
                min_value=min_r,
                max_value=max_r,
                value=(st.session_state.from_row, st.session_state.to_row)
            )
            st.session_state.from_row = r_range[0]
            st.session_state.to_row = r_range[1]
        
        st.divider()
        c1, c2 = st.columns(2)
        with c1:
            if st.button("🚀 Обробити / Оновити", type="primary"):
                sliced = slice_range(st.session_state.ld, st.session_state.from_row, st.session_state.to_row)
                st.session_state.sliced = sliced
                qinfo = classify_questions(sliced)
                st.session_state.qinfo = qinfo
                summaries = build_all_summaries(sliced, qinfo)
                st.session_state.summaries = summaries
                st.session_state.processed = True
        with c2:
            if st.button("❌ Скинути"):
                st.session_state.ld = None
                st.session_state.uploaded_files_store = None
                st.session_state.processed = False
                st.session_state.sliced = None
                st.session_state.summaries = None
                st.rerun()

# --- MAIN ---
if st.session_state.processed and st.session_state.sliced is not None:
    sliced = st.session_state.sliced
    summaries = st.session_state.summaries
    
    tab1, tab2 = st.tabs(["📊 Аналіз", "📥 Експорт"])
    
    with tab1:
        st.info(f"**В роботі {len(sliced)} анкет** (рядки {st.session_state.from_row}–{st.session_state.to_row})")
        
        with st.expander("🔍 Перегляд вихідних даних", expanded=False):
            st.dataframe(sliced, use_container_width=True)
        
        st.divider()
        st.subheader("Детальний перегляд")
        options = [qs.question.code for qs in summaries]
        selected_code = st.selectbox("Оберіть питання:", options)
        
        if selected_code:
            st.session_state.selected_code = selected_code
            selected = next((qs for qs in summaries if qs.question.code == st.session_state.selected_code), None)

            if selected and not selected.table.empty:
                st.markdown(f"**{selected.question.code}. {selected.question.text}**")
                c_ch, c_tb = st.columns([1.5, 1])
                with c_ch:
                    fig = px.pie(selected.table, names="Варіант відповіді", values="Кількість", hole=0, title="Розподіл")
                    st.plotly_chart(fig, use_container_width=True)
                with c_tb:
                    st.dataframe(selected.table, use_container_width=True)
            else:
                st.warning("Немає даних.")

        st.divider()
        st.subheader("📋 Повний огляд всіх питань")
        for qs in summaries:
            if qs.table.empty: continue
            with st.expander(f"{qs.question.code}. {qs.question.text}", expanded=True):
                c1, c2 = st.columns([1, 1])
                with c1:
                    st.plotly_chart(px.pie(qs.table, names="Варіант відповіді", values="Кількість", hole=0), use_container_width=True, key=f"ch_{qs.question.code}")
                with c2:
                    st.dataframe(qs.table, use_container_width=True)

    with tab2:
        st.subheader("Експорт")
        range_info = f"Рядки {st.session_state.from_row}–{st.session_state.to_row}"
        
        if os.path.exists("background.png"):
            st.success("✅ Фон 'background.png' знайдено.")
        
        @st.cache_data(show_spinner="Генеруємо Excel...")
        def get_excel_data(_original_df, _sliced_df, _qinfo, _summaries, _range_info):
            return build_excel_report(_original_df, _sliced_df, _qinfo, _summaries, _range_info)

        @st.cache_data(show_spinner="Генеруємо PDF...")
        def get_pdf_data(_original_df, _sliced_df, _summaries, _range_info):
            return build_pdf_report(_original_df, _sliced_df, _summaries, _range_info)

        @st.cache_data(show_spinner="Генеруємо DOCX...")
        def get_docx_data(_original_df, _sliced_df, _summaries, _range_info):
            return build_docx_report(_original_df, _sliced_df, _summaries, _range_info)

        @st.cache_data(show_spinner="Генеруємо PPTX...")
        def get_pptx_data(_original_df, _sliced_df, _summaries, _range_info):
            return build_pptx_report(_original_df, _sliced_df, _summaries, _range_info)

        cols = st.columns(4)
        with cols[0]:
            if st.button("📊 Excel"):
                data = get_excel_data(st.session_state.ld.df, st.session_state.sliced, st.session_state.qinfo, st.session_state.summaries, range_info)
                st.download_button("📥 Завантажити", data, "survey.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with cols[1]:
            if st.button("📄 PDF"):
                data = get_pdf_data(st.session_state.ld.df, st.session_state.sliced, st.session_state.summaries, range_info)
                st.download_button("📥 Завантажити", data, "survey.pdf", "application/pdf")
        with cols[2]:
            if st.button("📝 Word"):
                data = get_docx_data(st.session_state.ld.df, st.session_state.sliced, st.session_state.summaries, range_info)
                st.download_button("📥 Завантажити", data, "survey.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        with cols[3]:
            if st.button("🖥️ PPTX"):
                data = get_pptx_data(st.session_state.ld.df, st.session_state.sliced, st.session_state.summaries, range_info)
                st.download_button("📥 Завантажити", data, "survey.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation")

elif not st.session_state.ld:
    st.info("👈 Завантажте файл у меню ліворуч.")