"""
Модуль експорту звіту у формат PowerPoint (.pptx).
Оновлено: 
- Прибрано користувацькі шаблони.
- Додано автоматичне встановлення фонового зображення (background.png) для всіх слайдів.
- Стандартний заголовок.
"""

import io
import os
import textwrap
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

from classification import QuestionInfo, QuestionType
from summary import QuestionSummary
from typing import List, Optional

# --- КОНСТАНТИ ---
CHART_DPI = 150
FONT_SIZE_BASE = 10
BAR_WIDTH = 0.6

def create_chart_image(qs: QuestionSummary) -> io.BytesIO:
    """Генерує зображення діаграми для вставки у слайд."""
    plt.clf()
    plt.rcParams.update({'font.size': FONT_SIZE_BASE})
    
    labels = qs.table["Варіант відповіді"].astype(str).tolist()
    values = qs.table["Кількість"]
    wrapped_labels = [textwrap.fill(l, 30) for l in labels]

    if qs.question.qtype == QuestionType.SCALE:
        fig = plt.figure(figsize=(5.5, 4.0))
        bars = plt.bar(wrapped_labels, values, color='#4F81BD', width=BAR_WIDTH)
        plt.ylabel('Кількість')
        plt.grid(axis='y', linestyle='--', alpha=0.5)
        plt.xticks(rotation=0)
        for bar in bars:
            height = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                     f'{int(height)}', ha='center', va='bottom', fontweight='bold')
    else:
        fig = plt.figure(figsize=(5.5, 4.5))
        colors = ['#4F81BD', '#C0504D', '#9BBB59', '#8064A2', '#4BACC6', '#F79646']
        c_arg = colors[:len(values)] if len(values) <= len(colors) else None
        
        wedges, texts, autotexts = plt.pie(
            values, labels=None, autopct='%1.1f%%', startangle=90,
            pctdistance=0.8, colors=c_arg, radius=1.1,
            textprops={'fontsize': FONT_SIZE_BASE}
        )
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_weight('bold')
            import matplotlib.patheffects as path_effects
            autotext.set_path_effects([path_effects.withStroke(linewidth=2, foreground='#333333')])

        plt.axis('equal')
        cols = 2 if len(labels) > 2 else 1
        plt.legend(wrapped_labels, loc="upper center", bbox_to_anchor=(0.5, 0.0), ncol=cols, frameon=False, fontsize=9)

    plt.tight_layout()
    img_stream = io.BytesIO()
    plt.savefig(img_stream, format='png', dpi=CHART_DPI, bbox_inches='tight')
    plt.close(fig)
    img_stream.seek(0)
    return img_stream

def build_pptx_report(
    original_df: pd.DataFrame,
    sliced_df: pd.DataFrame,
    summaries: List[QuestionSummary],
    range_info: str,
    background_image_path: Optional[str] = "background.png" # Шлях до картинки
) -> bytes:
    
    prs = Presentation()

    # --- ВСТАНОВЛЕННЯ ФОНУ ---
    # Якщо файл існує, ставимо його фоном для Майстер-слайду (це вплине на всі макети)
    if background_image_path and os.path.exists(background_image_path):
        try:
            slide_master = prs.slide_master
            slide_master.background.fill.user_picture(background_image_path)
        except Exception as e:
            print(f"Не вдалося встановити фон: {e}")

    # 1. Титульний
    try: slide_layout = prs.slide_layouts[0]
    except: slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    
    try:
        title = slide.shapes.title
        title.text = "Звіт про результати опитування" # Стандартний заголовок
    except: pass
    
    try:
        if len(slide.placeholders) > 1:
            subtitle = slide.placeholders[1]
            subtitle.text = f"Всього анкет: {len(original_df)}\nОброблено: {len(sliced_df)}\n{range_info}"
    except: pass

    # 2. Технічний
    try:
        slide_layout = prs.slide_layouts[1] 
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        title.text = "Технічна інформація"
        body_shape = slide.shapes.placeholders[1]
        tf = body_shape.text_frame
        tf.text = "Параметри вибірки:"
        p = tf.add_paragraph(); p.text = f"Загальна кількість респондентів: {len(original_df)}"; p.level = 1
        p = tf.add_paragraph(); p.text = f"Кількість анкет у звіті: {len(sliced_df)}"; p.level = 1
        p = tf.add_paragraph(); p.text = f"Діапазон обробки: {range_info}"; p.level = 1
    except: pass

    # 3. Питання
    layout_index = 5 if len(prs.slide_layouts) > 5 else len(prs.slide_layouts) - 1
    
    for qs in summaries:
        if qs.table.empty: continue
        slide_layout = prs.slide_layouts[layout_index]
        slide = prs.slides.add_slide(slide_layout)
        try:
            title = slide.shapes.title
            title.text = f"{qs.question.code}. {qs.question.text}"
            if hasattr(title, 'text_frame'):
                 if len(title.text) > 60: title.text_frame.paragraphs[0].font.size = Pt(24)
                 else: title.text_frame.paragraphs[0].font.size = Pt(32)
        except: pass

        rows = len(qs.table) + 1; cols = 3
        left = Inches(0.5); top = Inches(2.0); width = Inches(4.5); height = Inches(0.8)
        table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
        table = table_shape.table
        table.columns[0].width = Inches(2.5); table.columns[1].width = Inches(1.0); table.columns[2].width = Inches(1.0)

        headers = ["Варіант", "Кільк.", "%"]
        for i, h in enumerate(headers):
            cell = table.cell(0, i)
            cell.text = h
            cell.text_frame.paragraphs[0].font.bold = True
            cell.text_frame.paragraphs[0].font.size = Pt(11)
            # Якщо фон темний, можливо, треба зробити таблицю напівпрозорою або білою
            # Тут ставимо білу заливку, щоб читалось на будь-якому фоні
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(255, 255, 255) 

        for i, row in enumerate(qs.table.itertuples(index=False)):
            cell = table.cell(i+1, 0); cell.text = str(row[0]); cell.text_frame.paragraphs[0].font.size = Pt(10)
            # Заливка даних теж біла, щоб не зливались з картинкою фону
            cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(255, 255, 255)
            
            cell = table.cell(i+1, 1); cell.text = str(row[1]); cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER; cell.text_frame.paragraphs[0].font.size = Pt(10)
            cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(255, 255, 255)
            
            cell = table.cell(i+1, 2); cell.text = str(row[2]); cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER; cell.text_frame.paragraphs[0].font.size = Pt(10)
            cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(255, 255, 255)

        try:
            img_stream = create_chart_image(qs)
            left_pic = Inches(5.2); top_pic = Inches(2.0); width_pic = Inches(4.5) 
            slide.shapes.add_picture(img_stream, left_pic, top_pic, width=width_pic)
        except: pass

    output = io.BytesIO()
    prs.save(output)
    return output.getvalue()