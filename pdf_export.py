import io
import textwrap
import pandas as pd # Обов'язково
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from fpdf import FPDF

from classification import QuestionInfo, QuestionType
from summary import QuestionSummary
from typing import List

CHART_DPI = 150
BAR_WIDTH = 0.6

class PDFReport(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 10)
        self.cell(0, 10, 'Звіт опитування', 0, 1, 'R')

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

def create_chart_image(qs: QuestionSummary) -> io.BytesIO:
    plt.close('all')
    plt.clf()
    plt.rcParams.update({'font.size': 10})
    
    labels = qs.table["Варіант відповіді"].astype(str).tolist()
    values = qs.table["Кількість"]
    wrapped_labels = [textwrap.fill(l, 25) for l in labels]

    # --- РОЗУМНА ПЕРЕВІРКА ТИПУ ---
    is_scale = (qs.question.qtype == QuestionType.SCALE)
    if not is_scale:
        try:
            vals = pd.to_numeric(qs.table["Варіант відповіді"], errors='coerce')
            if vals.notna().all() and vals.min() >= 0 and vals.max() <= 10:
                is_scale = True
        except: pass

    if is_scale:
        # СТОВПЧИКОВА
        fig = plt.figure(figsize=(6.0, 4.0))
        bars = plt.bar(wrapped_labels, values, color='#4F81BD', width=BAR_WIDTH)
        plt.ylabel('Кількість')
        plt.grid(axis='y', linestyle='--', alpha=0.5)
        for bar in bars:
            height = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                     f'{int(height)}', ha='center', va='bottom', fontweight='bold')
    else:
        # КРУГОВА
        fig = plt.figure(figsize=(6.0, 4.0))
        colors = ['#4F81BD', '#C0504D', '#9BBB59', '#8064A2', '#4BACC6', '#F79646']
        c_arg = colors[:len(values)] if len(values) <= len(colors) else None
        
        wedges, texts, autotexts = plt.pie(
            values, labels=None, autopct='%1.1f%%', startangle=90,
            pctdistance=0.8, colors=c_arg, radius=1.0
        )
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_weight('bold')
            import matplotlib.patheffects as path_effects
            autotext.set_path_effects([path_effects.withStroke(linewidth=2, foreground='#333333')])

        plt.axis('equal')
        cols = 2 if len(labels) > 3 else 1
        plt.legend(wrapped_labels, loc="upper center", bbox_to_anchor=(0.5, 0.0), ncol=cols, frameon=False, fontsize=8)

    plt.tight_layout()
    img_stream = io.BytesIO()
    plt.savefig(img_stream, format='png', dpi=CHART_DPI, bbox_inches='tight')
    plt.close(fig)
    img_stream.seek(0)
    return img_stream

def build_pdf_report(original_df, sliced_df, summaries, range_info) -> bytes:
    pdf = PDFReport()
    pdf.add_page()
    
    # Шрифти (Спроба кирилиці або fallback)
    try:
        pdf.add_font('DejaVu', '', 'DejaVuSans.ttf', uni=True)
        pdf.set_font('DejaVu', '', 12)
    except:
        pdf.set_font('Arial', '', 12)

    pdf.set_font_size(16)
    pdf.cell(0, 10, "Звіт про результати", 0, 1, 'C')
    pdf.set_font_size(12)
    pdf.cell(0, 10, f"Всього: {len(original_df)}, Оброблено: {len(sliced_df)}", 0, 1, 'C')
    pdf.cell(0, 10, range_info, 0, 1, 'C')
    pdf.ln(10)

    for qs in summaries:
        if qs.table.empty: continue
        
        title = f"{qs.question.code}. {qs.question.text}"
        pdf.set_font_size(12)
        pdf.multi_cell(0, 6, title)
        pdf.ln(2)

        # Table
        pdf.set_font_size(10)
        pdf.cell(100, 8, "Варіант", 1)
        pdf.cell(30, 8, "Кільк.", 1)
        pdf.cell(30, 8, "%", 1)
        pdf.ln()
        
        for row in qs.table.itertuples(index=False):
            val_text = str(row[0])[:50]
            pdf.cell(100, 8, val_text, 1)
            pdf.cell(30, 8, str(row[1]), 1)
            pdf.cell(30, 8, str(row[2]), 1)
            pdf.ln()
        pdf.ln(5)

        # Chart
        try:
            img = create_chart_image(qs)
            import tempfile, os
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                tmp.write(img.getvalue())
                tmp_path = tmp.name
            
            pdf.image(tmp_path, w=140)
            pdf.ln(10)
            os.unlink(tmp_path)
        except:
            pdf.cell(0, 10, "[Error chart]", 0, 1)

        if pdf.get_y() > 250: pdf.add_page()

    return pdf.output(dest='S').encode('latin-1', 'ignore')