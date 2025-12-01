"""
Модуль експорту звіту у форматі PDF.
Фінальні покращення:
- "Хірургічно точний" розрахунок висоти рядків таблиці (симуляція переносу слів).
- Компактні стовпчикові діаграми (висота 3.8 дюйма).
- Високі кругові діаграми з легендою знизу.
"""

import io
import os
import math
import textwrap
import tempfile
import requests
import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
from classification import QuestionInfo, QuestionType
from summary import QuestionSummary
from typing import List, Optional

# --- КОНСТАНТИ ---
FONT_URL = "https://github.com/dejavu-fonts/dejavu-fonts/raw/master/ttf/DejaVuSans.ttf"
FONT_NAME = "DejaVuSans"

CHART_DPI = 150           
FONT_SIZE_BASE = 11       
BAR_WIDTH = 0.6           

def get_font_path() -> Optional[str]:
    """Шукає або завантажує шрифт DejaVuSans.ttf."""
    local_path = "DejaVuSans.ttf"
    
    if os.path.exists(local_path) and os.path.getsize(local_path) > 10000:
        return local_path

    system_paths = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/TTF/DejaVuSans.ttf",
        "/usr/local/share/fonts/DejaVuSans.ttf"
    ]
    for path in system_paths:
        if os.path.exists(path):
            return path

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
        if self.get_y() > 240:
            self.add_page()
        self.set_font(FONT_NAME, "", 12)
        self.set_fill_color(220, 230, 241) 
        self.multi_cell(0, 8, str(text), fill=True, align='L')
        self.ln(2)

    def get_real_lines_count(self, text_val, col_width_mm):
        """
        Точний розрахунок кількості рядків, яку займе текст у комірці.
        Використовує get_string_width для симуляції переносу слів.
        """
        if not str(text_val):
            return 1
            
        # Ефективна ширина (ширина колонки мінус внутрішні відступи ~2-3мм)
        effective_width = col_width_mm - 4 
        
        # Розбиваємо на явні абзаци
        paragraphs = str(text_val).split('\n')
        total_lines = 0
        
        # Ширина пробілу
        space_w = self.get_string_width(' ')
        
        for p in paragraphs:
            if not p:
                total_lines += 1
                continue
                
            words = p.split(' ')
            current_line_w = 0
            lines_in_paragraph = 1
            
            for word in words:
                word_w = self.get_string_width(word)
                
                # Якщо слово довше за рядок (рідко, але буває), воно розірветься
                # Тут спрощено: вважаємо, що воно додається до рядка
                if current_line_w + word_w > effective_width:
                    # Перенос на новий рядок
                    lines_in_paragraph += 1
                    current_line_w = word_w + space_w
                else:
                    current_line_w += word_w + space_w
            
            total_lines += lines_in_paragraph
            
        return total_lines

    def add_table(self, df: pd.DataFrame):
        self.set_font(FONT_NAME, "", 10)
        line_height = 6
        col_width = [110, 30, 20] 
        headers = df.columns.tolist()

        # 1. Точний розрахунок висоти кожного рядка
        row_heights = []
        total_table_height = line_height # шапка
        
        for row in df.itertuples(index=False):
            text_val = str(row[0])
            # Використовуємо нову точну функцію
            n_lines = self.get_real_lines_count(text_val, col_width[0])
            h = n_lines * line_height
            row_heights.append(h)
            total_table_height += h

        # 2. Логіка переносу таблиці
        # Якщо таблиця середня (до 230мм) і не влазить -> переносимо всю
        page_limit = 275
        space_left = page_limit - self.get_y()
        
        if total_table_height < 230 and total_table_height > space_left:
            self.add_page()

        # 3. Друк
        self.set_fill_color(240, 240, 240)
        # Шапка
        for i, h in enumerate(headers):
            w = col_width[i] if i < len(col_width) else 20
            self.cell(w, line_height, str(h), border=1, fill=True, align='C')
        self.ln(line_height)

        # Рядки
        for idx, row in enumerate(df.itertuples(index=False)):
            text_val = str(row[0])
            count_val = str(row[1])
            perc_val = str(row[2])
            
            curr_h = row_heights[idx]

            # Перевірка: чи влізе цей КОНКРЕТНИЙ рядок?
            if self.get_y() + curr_h > page_limit:
                self.add_page()
                # Дублюємо шапку
                for i, h in enumerate(headers):
                    w = col_width[i] if i < len(col_width) else 20
                    self.cell(w, line_height, str(h), border=1, fill=True, align='C')
                self.ln(line_height)

            x_start = self.get_x()
            y_start = self.get_y()

            # Текст
            self.multi_cell(col_width[0], line_height, text_val, border=1, align='L')
            
            x_next = self.get_x()
            y_next = self.get_y()
            h_real = y_next - y_start 
            
            # Страховка: якщо раптом реальна висота більша за розрахункову
            final_h = max(h_real, curr_h)

            self.set_xy(x_start + col_width[0], y_start)
            self.cell(col_width[1], final_h, count_val, border=1, align='C')
            self.cell(col_width[2], final_h, perc_val, border=1, align='C')
            
            self.set_xy(x_start, y_start + final_h)

    def add_chart(self, qs: QuestionSummary):
        if qs.table.empty:
            return

        # Різні вимоги до місця
        space_needed = 110 if qs.question.qtype == QuestionType.SCALE else 170
        if self.get_y() > (280 - space_needed/2):
             self.add_page()

        plt.rcParams.update({'font.size': FONT_SIZE_BASE}) 
        
        labels = qs.table["Варіант відповіді"].astype(str).tolist()
        values = qs.table["Кількість"]
        wrapped_labels = [textwrap.fill(l, 40) for l in labels]

        if qs.question.qtype == QuestionType.SCALE:
            # --- СТОВПЧИКОВА: ЩЕ НИЖЧА ---
            # figsize=(10, 3.8) робить їх дуже акуратними
            plt.figure(figsize=(10, 3.8)) 
            
            bars = plt.bar(wrapped_labels, values, color='#4F81BD', width=BAR_WIDTH)
            plt.ylabel('Кількість')
            plt.grid(axis='y', linestyle='--', alpha=0.5)
            plt.xticks(rotation=0) 
            
            for bar in bars:
                height = bar.get_height()
                plt.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                         f'{int(height)}', ha='center', va='bottom', fontweight='bold')
        else:
            # --- КРУГОВА: ВИСОКА ---
            plt.figure(figsize=(10, 7))
            
            colors = ['#4F81BD', '#C0504D', '#9BBB59', '#8064A2', '#4BACC6', '#F79646']
            c_arg = colors[:len(values)] if len(values) <= len(colors) else None
            
            wedges, texts, autotexts = plt.pie(
                values, labels=None, autopct='%1.1f%%', startangle=90, 
                pctdistance=0.8, colors=c_arg, radius=1.2,
                textprops={'fontsize': FONT_SIZE_BASE}
            )
            
            for autotext in autotexts:
                autotext.set_color('white')
                autotext.set_weight('bold')
                import matplotlib.patheffects as path_effects
                autotext.set_path_effects([path_effects.withStroke(linewidth=2, foreground='#333333')])

            plt.axis('equal')
            
            cols = 2 if len(labels) > 3 else 1
            plt.legend(
                wrapped_labels, 
                loc="upper center", 
                bbox_to_anchor=(0.5, -0.05),
                ncol=cols,
                frameon=False
            )

        plt.tight_layout()

        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp_img:
            plt.savefig(tmp_img.name, format='png', dpi=CHART_DPI, bbox_inches='tight')
            tmp_img_path = tmp_img.name
        
        plt.close()

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
        err_pdf.multi_cell(0, 10, "CRITICAL ERROR: Cyrillic font not found.")
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
        
        try:
            pdf.add_chart(qs)
        except Exception:
            pass

        pdf.ln(5)

    return bytes(pdf.output())