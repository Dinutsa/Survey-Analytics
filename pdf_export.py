"""
Модуль експорту звіту у формат PDF.
Забезпечує гарантовану підтримку кирилиці (DejaVuSans), 
розумні розриви сторінок для таблиць та оптимізацію розміру діаграм.
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
    """Шукає або завантажує шрифт DejaVuSans.ttf."""
    local_path = "DejaVuSans.ttf"
    
    # Перевірка локального файлу
    if os.path.exists(local_path) and os.path.getsize(local_path) > 10000:
        return local_path

    # Перевірка системних шляхів
    system_paths = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/TTF/DejaVuSans.ttf",
        "/usr/local/share/fonts/DejaVuSans.ttf"
    ]
    for path in system_paths:
        if os.path.exists(path):
            return path

    # Завантаження
    try:
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
            raise RuntimeError("Font not found")
        self.add_font(FONT_NAME, "", font_path, uni=True)
        self.set_font(FONT_NAME, size=12)

    def header(self):
        self.set_font(FONT_NAME, "", 9)
        self.set_text_color(100, 100, 100)
        self.cell(0, 10, "Звіт за результатами опитування", border=False, align="R")
        self.ln(10)
        self.set_text_color(0, 0, 0)

    def footer(self):
        self.set_y(-15)
        self.set_font(FONT_NAME, "", 8)
        self.set_text_color(100, 100, 100)
        self.cell(0, 10, f"Сторінка {self.page_no()}", align="C")
        self.set_text_color(0, 0, 0)

    def chapter_title(self, text):
        # Якщо ми в самому низу сторінки, краще почати нову, щоб заголовок не висів сам
        if self.get_y() > 250:
            self.add_page()
            
        self.set_font(FONT_NAME, "", 12)
        self.set_fill_color(220, 230, 241) 
        self.multi_cell(0, 8, str(text), fill=True, align='L')
        self.ln(2)

    def draw_table_header(self, headers, col_width, line_height):
        """Допоміжна функція для малювання шапки таблиці."""
        self.set_fill_color(240, 240, 240)
        self.set_font(FONT_NAME, "", 10)
        x_start = self.get_x()
        y_start = self.get_y()
        
        # Малюємо шапку
        for i, h in enumerate(headers):
            w = col_width[i] if i < len(col_width) else 20
            self.cell(w, line_height, str(h), border=1, fill=True, align='C')
        self.ln(line_height)

    def add_table(self, df: pd.DataFrame):
        self.set_font(FONT_NAME, "", 10)
        line_height = 6  # Базова висота рядка
        col_width = [110, 30, 20] 
        headers = df.columns.tolist() 

        # Малюємо першу шапку
        self.draw_table_header(headers, col_width, line_height)

        for row in df.itertuples(index=False):
            text_val = str(row[0])
            count_val = str(row[1])
            perc_val = str(row[2])

            # 1. Розраховуємо висоту, яку займе цей рядок
            # FPDF не має прямого методу get_string_height для multi_cell до версії 2.5+
            # Тому ми робимо емуляцію: рахуємо кількість рядків тексту
            # Приблизно: довжина тексту / ширина колонки (в символах)
            # Але надійніше використовувати вбудований метод, якщо він є, або "на око" з запасом
            
            # Використовуємо multi_cell в режимі "dry run" або просто розрахунок рядків
            # В даному випадку просто перевіримо скільки рядків займе текст
            # Припускаємо, що ~60 символів влазить в 110мм при шрифті 10
            lines_count = max(1, len(text_val) // 55 + 1)
            # Додаємо трохи запасу на переноси слів
            if '\n' in text_val:
                lines_count += text_val.count('\n')
            
            row_height = lines_count * line_height

            # 2. Перевірка розриву сторінки
            # Якщо поточний Y + висота рядка > 275 (майже кінець А4), робимо нову сторінку
            if self.get_y() + row_height > 275:
                self.add_page()
                self.draw_table_header(headers, col_width, line_height)

            # 3. Малюємо рядок
            x_start = self.get_x()
            y_start = self.get_y()

            # Текст (Варіант)
            self.multi_cell(col_width[0], line_height, text_val, border=1, align='L')
            
            # Запам'ятовуємо, де закінчився multi_cell (реальна висота)
            x_next = self.get_x()
            y_next = self.get_y()
            h_real = y_next - y_start
            
            # Повертаємось наверх для малювання чисел
            self.set_xy(x_start + col_width[0], y_start)
            
            # Кількість
            self.cell(col_width[1], h_real, count_val, border=1, align='C')
            # %
            self.cell(col_width[2], h_real, perc_val, border=1, align='C')
            
            # Переходимо на наступний рядок (під найнижчу комірку)
            self.set_xy(x_start, y_next)

    def add_chart(self, qs: QuestionSummary):
        if qs.table.empty:
            return

        # Перевірка місця перед малюванням графіка
        # Графік займає десь 80-90 одиниць висоти
        if self.get_y() > 200:
            self.add_page()

        # Збільшуємо figsize для кращої якості, fpdf потім стисне до w=170
        plt.figure(figsize=(10, 5)) 
        
        labels = qs.table["Варіант відповіді"]
        values = qs.table["Кількість"]

        if qs.question.qtype == QuestionType.SCALE:
            # Bar chart
            bars = plt.bar(labels, values, color='#4F81BD', width=0.6)
            plt.ylabel('Кількість')
            plt.grid(axis='y', linestyle='--', alpha=0.5)
            # Значення над стовпчиками
            for bar in bars:
                height = bar.get_height()
                plt.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                         f'{int(height)}', ha='center', va='bottom', fontsize=10, fontweight='bold')
            plt.xticks(rotation=0)
        else:
            # Pie chart
            # startangle=90 виглядає звичніше
            colors = ['#4F81BD', '#C0504D', '#9BBB59', '#8064A2', '#4BACC6', '#F79646']
            # Тільки якщо даних менше ніж кольорів, інакше дефолтні
            c_arg = colors[:len(values)] if len(values) <= len(colors) else None
            
            plt.pie(values, labels=None, autopct='%1.1f%%', startangle=90, 
                    pctdistance=0.85, colors=c_arg, textprops={'fontsize': 10})
            
            # Легенда збоку
            plt.legend(labels, loc="center left", bbox_to_anchor=(1, 0.5), fontsize=10)
            plt.axis('equal') 

        plt.tight_layout()

        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp_img:
            # bbox_inches='tight' обрізає зайве біле поле!
            # pad_inches=0.1 додає маленький відступ, щоб підписи не різались
            plt.savefig(tmp_img.name, format='png', dpi=150, bbox_inches='tight', pad_inches=0.1)
            tmp_img_path = tmp_img.name
        
        plt.close()

        # Вставляємо на всю ширину сторінки (майже)
        # А4 ширина 210, відступи по 10 -> 190. Беремо 170 для краси.
        self.image(tmp_img_path, x=20, w=170)
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
    
    font_path = get_font_path()
    
    if not font_path:
        err_pdf = FPDF()
        err_pdf.add_page()
        err_pdf.set_font("Helvetica", size=12)
        err_pdf.multi_cell(0, 10, "CRITICAL ERROR: Cyrillic font (DejaVuSans) not found.\nCannot generate report.")
        return bytes(err_pdf.output())

    pdf = PDFReport(font_path)
    pdf.add_page()

    # --- Титульна ---
    pdf.set_font(FONT_NAME, "", 16)
    pdf.cell(0, 10, "Звіт про результати опитування", ln=True, align='C')
    pdf.ln(10)

    pdf.set_font(FONT_NAME, "", 12)
    pdf.cell(0, 8, f"Всього анкет: {len(original_df)}", ln=True)
    pdf.cell(0, 8, f"Оброблено анкет: {len(sliced_df)}", ln=True)
    pdf.cell(0, 8, f"Діапазон: {range_info}", ln=True)
    pdf.ln(10)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(10)

    # --- Основний цикл ---
    for qs in summaries:
        
        # Друкуємо заголовок
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
        
        # Додаємо діаграму
        try:
            pdf.add_chart(qs)
        except Exception:
            pass

        pdf.ln(5)

    return bytes(pdf.output())