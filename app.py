import io
import os
import zipfile
import streamlit as st
import plotly.express as px
import pandas as pd
import matplotlib.pyplot as plt

# Імпорти модулів
from data_loader import load_excels, get_row_bounds, slice_range
from classification import classify_questions, QuestionType
from summary import build_all_summaries

from excel_export import build_excel_report
from pdf_export import build_pdf_report
from docx_export import build_docx_report
from pptx_export import build_pptx_report

st.set_page_config(page_title="Обробка результатів", layout="wide")

# Ініціалізація стану
if 'processed' not in st.session_state: st.session_state.processed = False
if 'ld' not in st.session_state: st.session_state.ld = None
if 'uploaded_files_store' not in st.session_state: st.session_state.uploaded_files_store = None

st.title("Аналіз результатів опитувань (Google Forms)")

# --- SIDEBAR ---
with st.sidebar:
    st.header("1. Завантаження")
    uploaded_files = st.file_uploader("Excel-файли (.xlsx)", type=["xlsx"], accept_multiple_files=True)

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
            except Exception as e: st.error(f"Помилка: {e}")

    if st.session_state.ld:
        st.success(f"Завантажено: {st.session_state.ld.n_rows} анкет.")
        st.divider()
        st.header("2. Фільтрація")
        min_r, max_r = get_row_bounds(st.session_state.ld)
        if max_r > min_r:
            r_range = st.slider("Рядки", min_r, max_r, (st.session_state.from_row, st.session_state.to_row))
            st.session_state.from_row, st.session_state.to_row = r_range
        
        c1, c2 = st.columns(2)
        if c1.button("🚀 Обробити", type="primary"):
            sliced = slice_range(st.session_state.ld, st.session_state.from_row, st.session_state.to_row)
            st.session_state.sliced = sliced
            st.session_state.qinfo = classify_questions(sliced)
            st.session_state.summaries = build_all_summaries(sliced, st.session_state.qinfo)
            st.session_state.processed = True
            
        if c2.button("❌ Скинути"):
            st.session_state.clear()
            st.rerun()

# --- MAIN ---
if st.session_state.processed and st.session_state.sliced is not None:
    sliced = st.session_state.sliced
    summaries = st.session_state.summaries
    
    # Карта для пошуку: код -> об'єкт
    summary_map = {qs.question.code: qs for qs in summaries}
    question_codes = list(summary_map.keys())

    # Функція форматування для Selectbox
    def get_label(code):
        qs = summary_map[code]
        text = qs.question.text
        if len(text) > 90: text = text[:90] + "..."
        return f"{code}. {text}"

    t1, t2 = st.tabs(["📊 Аналіз", "📥 Експорт"])
    
    # === ВКЛАДКА 1: АНАЛІЗ ===
    with t1:
        st.info(f"**В роботі {len(sliced)} анкет** (рядки {st.session_state.from_row}–{st.session_state.to_row})")
        with st.expander("🔍 Перегляд вихідних даних", expanded=False): 
            st.dataframe(sliced, use_container_width=True)
        
        st.divider()
        
        # 1. ДЕТАЛЬНИЙ ПЕРЕГЛЯД
        st.subheader("Детальний перегляд")
        selected_code = st.selectbox("Оберіть питання:", options=question_codes, format_func=get_label, key="sb_detail")

        if selected_code:
            selected_qs = summary_map[selected_code]
            if not selected_qs.table.empty:
                st.markdown(f"**{selected_qs.question.text}**")
                c1, c2 = st.columns([1.5, 1])
                with c1: st.plotly_chart(px.pie(selected_qs.table, names="Варіант відповіді", values="Кількість", hole=0, title="Розподіл"), use_container_width=True)
                with c2: st.dataframe(selected_qs.table, use_container_width=True)
            else: st.warning("Немає даних.")

        st.divider()

        # 2. ПОДВІЙНА КРОС-ТАБУЛЯЦІЯ (НОВЕ!)
        st.subheader("🔀 Глибокий аналіз (Мульти-фільтр)")
        st.caption("Приклад: Як відповіли студенти **1 курсу** (Фільтр 1) про викладача **Петренка** (Фільтр 2)?")
        
        with st.expander("Налаштувати фільтри", expanded=True):
            
            # --- ФІЛЬТР 1 ---
            st.markdown("#### 1️⃣ Перший критерій")
            f1_col1, f1_col2 = st.columns(2)
            with f1_col1:
                filter1_code = st.selectbox("Питання:", options=question_codes, format_func=get_label, key="f1_q")
                filter1_qs = summary_map[filter1_code] if filter1_code else None
            with f1_col2:
                filter1_val = None
                if filter1_qs:
                    col1_name = filter1_qs.question.text
                    if col1_name in sliced.columns:
                        vals1 = [x for x in sliced[col1_name].unique() if pd.notna(x)]
                        filter1_val = st.selectbox("Значення:", vals1, key="f1_v")

            # --- ФІЛЬТР 2 (ОПЦІОНАЛЬНИЙ) ---
            use_filter2 = st.checkbox("➕ Додати другий критерій (звузити пошук)")
            filter2_qs = None
            filter2_val = None

            if use_filter2:
                st.markdown("#### 2️⃣ Другий критерій")
                f2_col1, f2_col2 = st.columns(2)
                with f2_col1:
                    filter2_code = st.selectbox("Питання:", options=question_codes, format_func=get_label, key="f2_q")
                    filter2_qs = summary_map[filter2_code] if filter2_code else None
                with f2_col2:
                    if filter2_qs:
                        col2_name = filter2_qs.question.text
                        if col2_name in sliced.columns:
                            # Тут хитрий момент: показуємо значення, які доступні ПІСЛЯ першого фільтру? 
                            # Або всі? Простіше показати всі, щоб не заплутати.
                            vals2 = [x for x in sliced[col2_name].unique() if pd.notna(x)]
                            filter2_val = st.selectbox("Значення:", vals2, key="f2_v")

            st.divider()

            # --- ЦІЛЬОВЕ ПИТАННЯ ---
            st.markdown("#### 🎯 Що аналізуємо?")
            target_code = st.selectbox("Питання для аналізу:", options=question_codes, format_func=get_label, key="target_q")
            target_qs = summary_map[target_code] if target_code else None

            # --- ЛОГІКА ФІЛЬТРАЦІЇ ---
            if st.button("🔍 Застосувати фільтри", type="primary"):
                if filter1_qs and filter1_val and target_qs:
                    
                    # 1. Застосовуємо Фільтр 1
                    subset = sliced[sliced[filter1_qs.question.text] == filter1_val]
                    info_text = f"Фільтр 1: {filter1_code} = '{filter1_val}'"

                    # 2. Якщо є, застосовуємо Фільтр 2
                    if use_filter2 and filter2_qs and filter2_val:
                        subset = subset[subset[filter2_qs.question.text] == filter2_val]
                        info_text += f" + Фільтр 2: {filter2_code} = '{filter2_val}'"

                    if not subset.empty:
                        st.success(f"Знайдено **{len(subset)}** анкет. ({info_text})")
                        
                        st.markdown(f"### Результат для: {target_qs.question.code}")
                        st.caption(target_qs.question.text)

                        # Статистика
                        col_target_name = target_qs.question.text
                        counts = subset[col_target_name].value_counts().reset_index()
                        counts.columns = ["Варіант відповіді", "Кількість"]
                        counts["%"] = (counts["Кількість"] / len(subset) * 100).round(1)
                        
                        # Графіки
                        g1, g2 = st.columns([1.5, 1])
                        with g1:
                            fig = px.pie(counts, names="Варіант відповіді", values="Кількість", hole=0, title="Розподіл")
                            st.plotly_chart(fig, use_container_width=True)
                        with g2:
                            st.dataframe(counts, use_container_width=True)
                    else:
                        st.error(f"Немає анкет, які відповідають обом умовам:\n1. {filter1_val}\n2. {filter2_val if use_filter2 else '-'}")
                else:
                    st.warning("Будь ласка, оберіть параметри фільтрації.")

        st.divider()
        
        # 3. ПОВНИЙ СПИСОК
        st.subheader("📋 Повний огляд всіх питань")
        for q in summaries:
            if q.table.empty: continue
            with st.expander(f"{q.question.code}. {q.question.text}", expanded=True):
                c1, c2 = st.columns([1, 1])
                with c1: st.plotly_chart(px.pie(q.table, names="Варіант відповіді", values="Кількість", hole=0), use_container_width=True, key=f"all_{q.question.code}")
                with c2: st.dataframe(q.table, use_container_width=True)

    # === ВКЛАДКА 2: ЕКСПОРТ ===
    with t2:
        st.subheader("Експорт")
        range_info = f"Рядки {st.session_state.from_row}–{st.session_state.to_row}"
        
        @st.cache_data(show_spinner="Excel...")
        def get_excel(_ld, _sl, _qi, _sm, _ri): return build_excel_report(_ld, _sl, _qi, _sm, _ri)
        @st.cache_data(show_spinner="PDF...")
        def get_pdf(_ld, _sl, _sm, _ri): return build_pdf_report(_ld, _sl, _sm, _ri)
        @st.cache_data(show_spinner="DOCX...")
        def get_docx(_ld, _sl, _sm, _ri): return build_docx_report(_ld, _sl, _sm, _ri)
        @st.cache_data(show_spinner="PPTX...")
        def get_pptx(_ld, _sl, _sm, _ri): return build_pptx_report(_ld, _sl, _sm, _ri)

        @st.cache_data(show_spinner="Архівуємо...")
        def get_zip_archive(_ld, _sl, _qi, _sm, _ri):
            plt.close('all') 
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.writestr("results.xlsx", build_excel_report(_ld, _sl, _qi, _sm, _ri))
                plt.close('all') 
                zf.writestr("results.pdf", build_pdf_report(_ld, _sl, _sm, _ri))
                plt.close('all') 
                zf.writestr("results.docx", build_docx_report(_ld, _sl, _sm, _ri))
                plt.close('all') 
                zf.writestr("results.pptx", build_pptx_report(_ld, _sl, _sm, _ri))
            return buf.getvalue()

        c1, c2, c3, c4 = st.columns(4)
        if c1.button("📊 Excel"): c1.download_button("📥", get_excel(st.session_state.ld.df, sliced, st.session_state.qinfo, summaries, range_info), "s.xlsx")
        if c2.button("📄 PDF"): c2.download_button("📥", get_pdf(st.session_state.ld.df, sliced, summaries, range_info), "s.pdf")
        if c3.button("📝 Word"): c3.download_button("📥", get_docx(st.session_state.ld.df, sliced, summaries, range_info), "s.docx")
        if c4.button("🖥️ PPTX"): c4.download_button("📥", get_pptx(st.session_state.ld.df, sliced, summaries, range_info), "s.pptx")

        st.divider()
        if st.button("🗂️ Сформувати ZIP-архів", type="primary", use_container_width=True):
            zip_data = get_zip_archive(st.session_state.ld.df, sliced, st.session_state.qinfo, summaries, range_info)
            st.download_button("📥 Скачати ZIP", zip_data, "full_report.zip", "application/zip", type="primary", use_container_width=True)

elif not st.session_state.ld:
    st.info("👈 Завантажте файл.")