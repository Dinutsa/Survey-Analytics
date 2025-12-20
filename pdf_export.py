"""
Модуль експорту звіту у формат PDF.
ВЕРСІЯ: TIMES NEW ROMAN STYLE (Tinos Font).
- Використовує шрифт Tinos (аналог Times New Roman).
- Повна підтримка кирилиці.
- Розумні графіки (Стовпчики/Круг).
"""

import io
import os
import urllib.request
import textwrap
import tempfile
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

# Налаштування шрифту (Tinos - це 100% аналог Times New Roman з кирилицею)
FONT_FILENAME = "Tinos-Regular.ttf"
FONT_PATH = os.path.join(os.getcwd(), FONT_FILENAME)
# Пряме посилання на файл шрифту
FONT_URL = "https://github.com/google/fonts/raw/main/apache/tinos/Tinos-Regular.ttf"

def ensure_font_exists():
    """Гарантує, що файл шрифту є на диску."""
    if not os.path.exists(FONT_PATH) or os.path.getsize(FONT_PATH) == 0:
        try:
            print(f"🔄 Завантажую шрифт (Times style): {FONT_PATH}")
            opener = urllib.request.build_opener()
            opener.addheaders = [('User-agent', 'Mozilla/5.0')]
            urllib.request.install_opener(opener)
            urllib.request.urlretrieve(FONT_URL, FONT_PATH)
            print("✅ Шрифт успішно завантажено!")
        except Exception as e:
            print(f"❌ Помилка завантаження шрифту: {e}")

class PDFReport(FPDF):
    def header(self):
        # Використовуємо шрифт Times
        try:
            self.set_font("TimesUA", size=10)
            self.cell(0, 10, "Звіт про результати опитування", ln=1, align='R')
        except Exception:
            # Fallback
            self.set_font("Times", "B", 10)
            self.cell(0, 10, "Survey Report", ln=1, align='R')

    def footer(self):
        self.set_y(-15)
        try:
            self.set_font("TimesUA", size=8)
        except:
            self.set_font("Times", "I", 8)
        self.cell(0, 10, f'Page {self.page_no()}', align='C')

def create_chart_image(qs: QuestionSummary) -> io.BytesIO:
    plt.close('all')
    plt.clf()
    # Шрифт на графіках теж робимо схожим на Times (serif)
    plt.rcParams.update({
        'font.size': 10,
        'font.family': 'serif' 
    })
    
    labels = qs.table["Варіант відповіді"].astype(str).tolist()
    values = qs.table["Кількість"]
    wrapped_labels = [textwrap.fill(l, 25) for l in labels]

    is_scale = (qs.question.qtype == QuestionType.SCALE)
    if not is_scale:
        try:
            vals = pd.to_numeric(qs.table["Варіант відповіді"], errors='coerce')
            if vals.notna().all() and vals.min() >= 0 and vals.max() <= 10:
                is_scale = True
        except: pass

    if is_scale:
        fig = plt.figure(figsize=(6.0, 4.0))
        # Колір стовпчиків - класичний синій
        bars = plt.bar(wrapped_labels, values, color='#4F81BD', width=BAR_WIDTH)
        plt.ylabel('Кількість')
        plt.grid(axis='y', linestyle='--', alpha=0.5)
        for bar in bars:
            height = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                     f'{int(height)}', ha='center', va='bottom', fontweight='bold')
    else:
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
    ensure_font_exists()
    
    pdf = PDFReport()
    
    font_ok = False
    if os.path.exists(FONT_PATH) and os.path.getsize(FONT_PATH) > 0:
        try:
            # Реєструємо шрифт під назвою "TimesUA"
            pdf.add_font("TimesUA", fname=FONT_PATH)
            font_ok = True
        except Exception as e:
            print(f"⚠️ Font error: {e}")

    pdf.add_page()
    
    # --- ТИТУЛКА (Times New Roman Style) ---
    if font_ok:
        pdf.set_font("TimesUA", size=16)
        pdf.cell(0, 10, "Звіт про результати опитування", ln=1, align='C')
        
        pdf.set_font("TimesUA", size=12)
        safe_range = range_info.replace('–', '-').replace('—', '-')
        
        pdf.cell(0, 10, f"Всього анкет: {len(original_df)}", ln=1, align='C')
        pdf.cell(0, 10, f"Оброблено: {len(sliced_df)}", ln=1, align='C')
        pdf.cell(0, 10, safe_range, ln=1, align='C')
    else:
        # Fallback (якщо раптом шрифт не скачався)
        pdf.set_font("Times", "B", 16)
        pdf.cell(0, 10, "Survey Report", ln=1, align='C')
        pdf.set_font("Times", size=12)
        pdf.cell(0, 10, f"Count: {len(sliced_df)}", ln=1, align='C')
    
    pdf.ln(5)

    for qs in summaries:
        if qs.table.empty: continue
        
        title = f"{qs.question.code}. {qs.question.text}"
        title = title.replace('–', '-').replace('—', '-').replace('’', "'")
        
        # Назва питання
        if font_ok:
            pdf.set_font("TimesUA", size=12) # Звичайний Times
            pdf.multi_cell(0, 6, title)
        else:
            pdf.set_font("Times", "B", 12)
            pdf.multi_cell(0, 6, f"Question {qs.question.code}")

        pdf.ln(2)

        # Таблиця
        if font_ok: pdf.set_font("TimesUA", size=11)
        else: pdf.set_font("Times", size=10)

        col_w1 = 110
        col_w2 = 30
        
        # Заголовки таблиці
        h1 = "Варіант" if font_ok else "Option"
        h2 = "Кільк." if font_ok else "Count"
        h3 = "%"
        
        pdf.cell(col_w1, 8, h1, border=1, ln=0)
        pdf.cell(col_w2, 8, h2, border=1, ln=0)
        pdf.cell(col_w2, 8, h3, border=1, ln=1)
        
        for row in qs.table.itertuples(index=False):
            val_text = str(row[0])[:60].replace('–', '-').replace('—', '-').replace('’', "'")
            
            # Якщо шрифт не завантажився, уникаємо кирилиці
            if not font_ok and not val_text.isascii():
                val_text = "..."

            pdf.cell(col_w1, 8, val_text, border=1, ln=0)
            pdf.cell(col_w2, 8, str(row[1]), border=1, ln=0)
            pdf.cell(col_w2, 8, str(row[2]), border=1, ln=1)
            
        pdf.ln(5)

        # Графік
        try:
            img = create_chart_image(qs)
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                tmp.write(img.getvalue())
                name = tmp.name
            
            pdf.image(name, w=140, x=35)
            os.unlink(name)
            pdf.ln(10)
        except:
            pdf.cell(0, 10, "[Chart Error]", ln=1)

        if pdf.get_y() > 240:
            pdf.add_page()

    # Повертаємо PDF
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
        pdf.output(tmp_pdf.name)
        tmp_name = tmp_pdf.name
        
    with open(tmp_name, 'rb') as f:
        pdf_bytes = f.read()
    os.unlink(tmp_name)
    
    return pdf_bytes