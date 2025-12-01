"""
Модуль експорту звіту у формат PDF.
Забезпечує підтримку кирилиці та генерацію статичних графіків.
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
from typing import List, Dict

# URL для завантаження шрифту з підтримкою кирилиці (DejaVuSans)
FONT_URL = "https://github.com/google/fonts/raw/main/ofl/notosans/NotoSans-Regular.ttf"
FONT_NAME = "NotoSans"

class PDFReport(FPDF):
    def __init__(self, font_path):
        super().__init__()
        self.font_path = font_path
        # Реєструємо шрифт
        self.add_font(FONT_NAME, "", font_path, uni=True)
        self.set_font(FONT_NAME, size=12)

    def header(self):
        # Заголовок на кожній сторінці
        self.set_font(FONT_NAME, "", 10)
        self.cell(0, 10, "Звіт за результатами опитування", border=False, align="R")
        self.ln(10)

    def footer(self):
        # Нумерація сторінок
        self.set_y(-15)
        self.set_font(FONT_NAME, "", 8)
        self.cell(0, 10, f"Сторінка {self.page_no()}", align="C")

    def chapter_title(self, text):
        self.set_font(FONT_NAME, "", 12)
        # Сірий фон для заголовка
        self.set_fill_color(220, 230, 241) 
        self.multi_cell(0, 10, text, fill=True, align='L')
        self.ln(2)

    def add_table(self, df: pd.DataFrame):
        """Малює таблицю з DataFrame."""
        self.set_font(FONT_NAME, "", 10)
        
        # Висота рядка
        line_height = self.font_size * 2
        col_width = [110, 30, 20]  # Ширина колонок: Варіант, Кількість, %

        # Заголовки
        headers = df.columns.tolist() # Варіант, Кількість, %
        self.set_fill_color(240, 240, 240)
        
        # Малюємо заголовок
        for i, h in enumerate(headers):
            w = col_width[i] if i < len(col_width) else 20
            self.cell(w, line_height, str(h), border=1, fill=True, align='C')
        self.ln(line_height)

        # Дані
        for row in df.itertuples(index=False):
            # row[0] - текст, row[1] - число, row[2] - %
            
            # Обробка довгого тексту у першій колонці
            x_start = self.get_x()
            y_start = self.get_y()
            
            text_val = str(row[0])
            count_val = str(row[1])
            perc_val = str(row[2])

            # Визначаємо висоту комірки на основі тексту
            # multi_cell не повертає висоту, тому робимо емуляцію або фіксуємо
            # Тут простий варіант: multi_cell для тексту, cell для чисел
            
            # 1. Текст (Варіант)
            self.multi_cell(col_width[0], line_height, text_val, border=1, align='L')
            
            # Позиція після multi_cell
            x_next = self.get_x()
            y_next = self.get_y()
            h_curr = y_next - y_start
            
            # Повертаємось вгору праворуч для чисел
            self.set_xy(x_start + col_width[0], y_start)
            
            # 2. Кількість
            self.cell(col_width[1], h_curr, count_val, border=1, align='C')
            
            # 3. Відсоток
            self.cell(col_width[2], h_curr, perc_val, border=1, align='C')
            
            self.ln()

    def add_chart(self, qs: QuestionSummary):
        """Створює графік через Matplotlib і вставляє в PDF."""
        # Якщо даних немає, пропускаємо
        if qs.table.empty:
            return

        plt.figure(figsize=(6, 3))
        
        # Підготовка даних
        labels = qs.table["Варіант відповіді"]
        values = qs.table["Кількість"]

        if qs.question.qtype == QuestionType.SCALE:
            # Стовпчикова
            bars = plt.bar(labels, values, color='#4F81BD')
            plt.ylabel('Кількість')
            # Додаємо значення над стовпчиками
            for bar in bars:
                height = bar.get_height()
                plt.text(bar.get_x() + bar.get_width()/2., height,
                         f'{int(height)}', ha='center', va='bottom')
            plt.xticks(rotation=0) # або 45, якщо довгі підписи
            
        else:
            # Кругова
            # Для краси відсікаємо дуже малі сегменти для підписів
            plt.pie(values, labels=None, autopct='%1.1f%%', startangle=140, pctdistance=0.85)
            plt.legend(labels, loc="center left", bbox_to_anchor=(1, 0.5), fontsize='small')
            plt.axis('equal') 

        plt.tight_layout()

        # Зберігаємо у тимчасовий файл
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp_img:
            plt.savefig(tmp_img.name, format='png', dpi=100)
            tmp_img_path = tmp_img.name
        
        plt.close() # Закриваємо фігуру, щоб звільнити пам'ять

        # Вставляємо в PDF
        # Перевіряємо, чи вистачить місця, якщо ні - нова сторінка
        if self.get_y() > 200:
            self.add_page()
            
        self.image(tmp_img_path, w=150)
        self.ln(5)
        
        # Видаляємо файл
        try:
            os.remove(tmp_img_path)
        except:
            pass

def ensure_font_exists():
    """Завантажує шрифт NotoSans/DejaVu, якщо його немає, для підтримки кирилиці."""
    temp_dir = tempfile.gettempdir()
    font_path = os.path.join(temp_dir, "NotoSans-Regular.ttf")
    
    if not os.path.exists(font_path):
        try:
            response = requests.get(FONT_URL)
            with open(font_path, "wb") as f:
                f.write(response.content)
        except Exception:
            # Якщо завантажити не вдалося, повертаємо None (буде помилка кирилиці, але код не впаде)
            return None
    return font_path

def build_pdf_report(
    original_df: pd.DataFrame,
    sliced_df: pd.DataFrame,
    summaries: List[QuestionSummary],
    range_info: str
) -> bytes:
    
    font_path = ensure_font_exists()
    if not font_path:
        # Fallback, якщо немає інтернету, спробуємо стандартний Arial (може не працювати на Linux)
        font_path = "Arial.ttf" 

    pdf = PDFReport(font_path)
    pdf.add_page()

    # --- Титульна сторінка / Технічна інформація ---
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

    # --- Основний цикл ---
    for qs in summaries:
        # Перевірка на розрив сторінки, якщо мало місця
        if pdf.get_y() > 250:
            pdf.add_page()

        # Заголовок питання
        title = f"{qs.question.code}. {qs.question.text}"
        pdf.chapter_title(title)

        if qs.table.empty:
            pdf.set_font(FONT_NAME, "", 10)
            pdf.cell(0, 10, "Немає даних або відкриті відповіді (текст).", ln=True)
            pdf.ln(5)
            continue

        # Таблиця
        try:
            pdf.add_table(qs.table)
        except Exception as e:
            pdf.cell(0, 10, "Помилка відображення таблиці", ln=True)

        # Діаграма
        pdf.ln(5)
        # Перевірка місця для діаграми
        if pdf.get_y() > 180:
            pdf.add_page()
            
        try:
            pdf.add_chart(qs)
        except Exception as e:
            print(f"Chart error: {e}")

        pdf.ln(10)

    # Повертаємо байти
    return pdf.output(dest='S').encode('latin-1')