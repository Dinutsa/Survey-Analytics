"""
Модуль експорту звіту у формат PDF.
ВЕРСІЯ: Universal Syntax (Legacy Compatible).
- Працює і з fpdf, і з fpdf2.
- Використовує ln=1 замість new_x/new_y.
- Автоматичне завантаження українського шрифту.
"""

import io
import os
import urllib.request
import textwrap
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from fpdf import FPDF

from classification import QuestionType
from summary import QuestionSummary

# --- НАЛАШТУВАННЯ ---
CHART_DPI = 150
BAR_WIDTH = 0.6
FONT_URL = "https://github.com/coreybutler/fonts/raw/master/ttf/DejaVuSans.ttf"
FONT_FILE = "DejaVuSans.ttf"

def check_and_download_font():
    """Завантажує шрифт, якщо його немає."""
    if not os.path.exists(FONT_FILE):
        try:
            print(f"Завантаження шрифту {FONT_FILE}...")
            opener = urllib.request.build_opener()
            opener.addheaders = [('User-agent', 'Mozilla/5.0')]
            urllib.request.install_opener(opener)
            urllib.request.urlretrieve(FONT_URL, FONT_FILE)
            print("Шрифт завантажено.")
        except Exception as e:
            print(f"Помилка завантаження шрифту: {e}")

class PDFReport(FPDF):
    def header(self):
        # Спроба встановити шрифт (DejaVu або стандартний)
        try:
            self.set_font("DejaVu", size=10)
        except:
            self.set_font("Arial", "B", 10)
        
        # Використовуємо ln=1 (новий рядок) замість new_y="NEXT"
        self.cell(0, 10, "Звіт про результати опитування", ln=1, align='R')

    def footer(self):
        self.set_y(-15)
        try:
            self.set_font("DejaVu", size=8)
        except:
            self.set_font("Arial", "I", 8)
        self.cell(0, 10, f'Page {self.page_no()}', align='C')

def create_chart_image(qs: QuestionSummary) -> io.BytesIO:
    plt.close('all')
    plt.clf()
    plt.rcParams.update({'font.size': 10})
    
    labels = qs.table["Варіант відповіді"].astype(str).tolist()
    values = qs.table["Кількість"]
    wrapped_labels = [textwrap.fill(l, 25) for l in labels]

    # Розумна перевірка типу (числа -> стовпчики)
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
    # 1. Завантажуємо шрифт
    check_and_download_font()
    
    pdf = PDFReport()
    
    # 2. Реєструємо шрифт (пробуємо різні варіанти параметрів для сумісності)
    font_ok = False
    if os.path.exists(FONT_FILE):
        try:
            # Універсальний виклик, який часто спрацьовує і там, і там
            pdf.add_font('DejaVu', '', FONT_FILE, uni=True)
            font_ok = True
        except:
            try:
                # Спроба для новіших версій без uni=True
                pdf.add_font('DejaVu', '', FONT_FILE)
                font_ok = True
            except:
                print("Не вдалося підключити шрифт DejaVu.")

    pdf.add_page()
    
    # Вибір шрифту
    if font_ok: pdf.set_font("DejaVu", size=16)
    else: pdf.set_font("Arial", "B", 16)
    
    pdf.cell(0, 10, "Звіт про результати", ln=1, align='C')
    
    if font_ok: pdf.set_font("DejaVu", size=12)
    else: pdf.set_font("Arial", size=12)

    pdf.cell(0, 10, f"Всього: {len(original_df)} | Оброблено: {len(sliced_df)}", ln=1, align='C')
    
    safe_range = range_info.replace('–', '-').replace('—', '-')
    pdf.cell(0, 10, safe_range, ln=1, align='C')
    pdf.ln(5)

    for qs in summaries:
        if qs.table.empty: continue
        
        # Заголовок
        title = f"{qs.question.code}. {qs.question.text}"
        title = title.replace('–', '-').replace('—', '-').replace('’', "'")
        
        if font_ok: pdf.set_font("DejaVu", size=12)
        else: pdf.set_font("Arial", size=12)
            
        pdf.multi_cell(0, 6, title)
        pdf.ln(2)

        # Таблиця
        if font_ok: pdf.set_font("DejaVu", size=10)
        else: pdf.set_font("Arial", size=10)

        col_w1 = 110
        col_w2 = 30
        
        # Заголовки (ln=0 означає "залишаємось на рядку")
        pdf.cell(col_w1, 8, "Варіант", border=1, ln=0)
        pdf.cell(col_w2, 8, "Кільк.", border=1, ln=0)
        pdf.cell(col_w2, 8, "%", border=1, ln=1) # Тут ln=1, перехід на новий рядок
        
        for row in qs.table.itertuples(index=False):
            val_text = str(row[0])[:60].replace('–', '-').replace('—', '-').replace('’', "'")
            
            pdf.cell(col_w1, 8, val_text, border=1, ln=0)
            pdf.cell(col_w2, 8, str(row[1]), border=1, ln=0)
            pdf.cell(col_w2, 8, str(row[2]), border=1, ln=1)
            
        pdf.ln(5)

        # Графік
        try:
            img = create_chart_image(qs)
            import tempfile
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                tmp.write(img.getvalue())
                name = tmp.name
            
            # x=35 центрує картинку (при ширині А4 ~210мм і ширині картинки 140мм)
            pdf.image(name, w=140, x=35)
            os.unlink(name)
            pdf.ln(10)
        except:
            pdf.cell(0, 10, "[Chart Error]", ln=1)

        if pdf.get_y() > 240:
            pdf.add_page()

    # Повертаємо байти (через тимчасовий файл, щоб уникнути помилок кодування)
    import tempfile
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
        pdf.output(tmp_pdf.name)
        tmp_name = tmp_pdf.name
        
    with open(tmp_name, 'rb') as f:
        pdf_bytes = f.read()
    os.unlink(tmp_name)
    
    return pdf_bytes