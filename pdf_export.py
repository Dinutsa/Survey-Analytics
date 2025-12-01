"""
Модуль експорту звіту у формат PDF.
Забезпечує гарантовану підтримку кирилиці (DejaVuSans).
"""

import io
import os
import tempfile
import requests
import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
from classification import QuestionInfo, QuestionType
from summary import QuestionSummary
from typing import List, Optional

# Посилання на шрифт
FONT_URL = "https://github.com/dejavu-fonts/dejavu-fonts/raw/master/ttf/DejaVuSans.ttf"
FONT_NAME = "DejaVuSans"

def get_font_path() -> Optional[str]:
    """
    Шукає шрифт з підтримкою кирилиці:
    1. У поточній папці.
    2. У системних папках Linux (Streamlit Cloud).
    3. Намагається завантажити з Інтернету.
    """
    local_path = "DejaVuSans.ttf"
    
    # 1. Перевіряємо локальний файл
    if os.path.exists(local_path) and os.path.getsize(local_path) > 10000:
        return local_path

    # 2. Перевіряємо системні шляхи (стандартні для Linux/Debian/Streamlit Cloud)
    system_paths = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/TTF/DejaVuSans.ttf",
        "/usr/local/share/fonts/DejaVuSans.ttf"
    ]
    for path in system_paths:
        if os.path.exists(path):
            return path

    # 3. Завантажуємо, якщо не знайшли
    try:
        print("Downloading font...")
        headers = {'User-Agent': 'Mozilla/5.0'}
        response = requests.get(FONT_URL, headers=headers, timeout=10)
        response.raise_for_status()
        with open(local_path, "wb") as f:
            f.write(response.content)
        return local_path
    except Exception as e:
        print(f"Font download failed: {e}")
        return None

class PDFReport(FPDF):
    def __init__(self, font_path):
        super().__init__()
        self.font_path = font_path
        
        if not font_path:
            raise RuntimeError("Критична помилка: Не знайдено шрифт для кирилиці.")

        # Реєструємо шрифт. uni=True важливий для fpdf < 2.5, для нових версій це стандарт
        self.add_font(FONT_NAME, "", font_path, uni=True)
        self.set_font(FONT_NAME, size=12)

    def header(self):
        self.set_font(FONT_NAME, "", 10)
        self.cell(0, 10, "Звіт за результатами опитування", border=False, align="R")
        self.ln(10)

    def footer(self):
        self.set_y(-15)
        self.set_font(FONT_NAME, "", 8)
        self.cell(0, 10, f"Сторінка {self.page_no()}", align="C")

    def chapter_title(self, text):
        self.set_font(FONT_NAME, "", 12)
        self.set_fill_color(220, 230, 241) 
        # text=str(text) для безпеки
        self.multi_cell(0, 10, str(text), fill=True, align='L')
        self.ln(2)

    def add_table(self, df: pd.DataFrame):
        self.set_font(FONT_NAME, "", 10)
        line_height = self.font_size * 2
        col_width = [110, 30, 20] 

        headers = df.columns.tolist() 
        self.set_fill_color(240, 240, 240)
        
        # Заголовок
        for i, h in enumerate(headers):
            w = col_width[i] if i < len(col_width) else 20
            self.cell(w, line_height, str(h), border=1, fill=True, align='C')
        self.ln(line_height)

        # Дані
        for row in df.itertuples(index=False):
            x_start = self.get_x()
            y_start = self.get_y()
            
            text_val = str(row[0])
            count_val = str(row[1])
            perc_val = str(row[2])

            # Текст (багаторядковий)
            self.multi_cell(col_width[0], line_height, text_val, border=1, align='L')
            
            x_next = self.get_x()
            y_next = self.get_y()
            h_curr = y_next - y_start
            
            # Числа (однорядкові, але висота підганяється під текст)
            self.set_xy(x_start + col_width[0], y_start)
            self.cell(col_width[1], h_curr, count_val, border=1, align='C')
            self.cell(col_width[2], h_curr, perc_val, border=1, align='C')
            
            # Повертаємось на початок нового рядка
            self.set_xy(x_start, y_next)

    def add_chart(self, qs: QuestionSummary):
        if qs.table.empty:
            return

        plt.figure(figsize=(6, 3))
        labels = qs.table["Варіант відповіді"]
        values = qs.table["Кількість"]

        if qs.question.qtype == QuestionType.SCALE:
            bars = plt.bar(labels, values, color='#4F81BD')
            plt.ylabel('Кількість')
            for bar in bars:
                height = bar.get_height()
                plt.text(bar.get_x() + bar.get_width()/2., height,
                         f'{int(height)}', ha='center', va='bottom')
            plt.xticks(rotation=0)
        else:
            plt.pie(values, labels=None, autopct='%1.1f%%', startangle=140, pctdistance=0.85)
            plt.legend(labels, loc="center left", bbox_to_anchor=(1, 0.5), fontsize='small')
            plt.axis('equal') 

        plt.tight_layout()

        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp_img:
            plt.savefig(tmp_img.name, format='png', dpi=100)
            tmp_img_path = tmp_img.name
        
        plt.close()

        if self.get_y() > 200:
            self.add_page()
            
        self.image(tmp_img_path, w=150)
        self.ln(5)
        
        try:
            os.remove(tmp_img_path)
        except:
            pass

def build_pdf_report(
    original_df: pd.DataFrame,
    sliced_df: pd.DataFrame,
    summaries: List[QuestionSummary],
    range_info: str
) -> bytes:
    
    # 1. Отримуємо шлях до шрифту
    font_path = get_font_path()
    
    # Якщо шрифт не знайдено, ми не можемо генерувати PDF з кирилицею.
    # Створюємо PDF з помилкою (англійською, щоб не впало).
    if not font_path:
        err_pdf = FPDF()
        err_pdf.add_page()
        err_pdf.set_font("Helvetica", size=12)
        err_pdf.multi_cell(0, 10, "CRITICAL ERROR: Cyrillic font (DejaVuSans) not found on server.\nCannot generate report.")
        return bytes(err_pdf.output())

    # 2. Генеруємо нормальний звіт
    pdf = PDFReport(font_path)
    pdf.add_page()

    # Титулка
    pdf.set_font(FONT_NAME, "", 16)
    pdf.cell(0, 10, "Звіт про результати опитування", ln=True, align='C')
    pdf.ln(10)

    pdf.set_font(FONT_NAME, "", 12)
    pdf.cell(0, 10, f"Всього анкет: {len(original_df)}", ln=True)
    pdf.cell(0, 10, f"Оброблено анкет: {len(sliced_df)}", ln=True)
    pdf.cell(0, 10, f"Діапазон: {range_info}", ln=True)
    pdf.ln(10)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(10)

    # Цикл по питаннях
    for qs in summaries:
        if pdf.get_y() > 250:
            pdf.add_page()

        title = f"{qs.question.code}. {qs.question.text}"
        pdf.chapter_title(title)

        if qs.table.empty:
            pdf.set_font(FONT_NAME, "", 10)
            pdf.cell(0, 10, "Немає даних або відкриті відповіді.", ln=True)
            pdf.ln(5)
            continue

        try:
            pdf.add_table(qs.table)
        except Exception as e:
            pdf.cell(0, 10, f"Table Error: {e}", ln=True)

        pdf.ln(5)
        
        if pdf.get_y() > 180:
            pdf.add_page()
            
        try:
            pdf.add_chart(qs)
        except Exception:
            pass

        pdf.ln(10)

    return bytes(pdf.output())