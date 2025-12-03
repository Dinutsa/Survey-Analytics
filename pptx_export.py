"""
Модуль експорту звіту у формат PowerPoint (.pptx).
ВЕРСІЯ: Native Style (Нативний стиль).
- Використовує вбудований стиль PowerPoint "Table Grid" (ID: {5940675A-B579-460E-94D1-54222C63F5DA}).
- Це гарантує наявність чорних меж (рамок) без складного коду.
- Фон заголовка фарбується вручну.
- Фон слайдів (background.png) накладається на всі макети.
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

# Імпорти для застосування стилю таблиці
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn

from classification import QuestionInfo, QuestionType
from summary import QuestionSummary
from typing import List, Optional

# --- НАЛАШТУВАННЯ ---
CHART_DPI = 150
FONT_SIZE_CHART = 12
FONT_SIZE_TABLE_HEADER = 12
FONT_SIZE_TABLE_DATA = 11
BAR_WIDTH = 0.6

# --- ФУНКЦІЯ ДЛЯ СТИЛЮ ТАБЛИЦІ ---
def apply_table_grid_style(table):
    """
    Застосовує стиль 'Table Grid' (Сітка таблиці) через XML.
    Це додає чорні рамки до всіх клітинок автоматично.
    """
    tbl = table._tbl
    tblPr = tbl.tblPr
    
    # Шукаємо або створюємо посилання на стиль
    tblStyle = tblPr.find(qn('a:tableStyleId'))
    if tblStyle is None:
        tblStyle = OxmlElement('a:tableStyleId')
        tblPr.append(tblStyle)
    
    # ID стилю "Table Grid" (стандартний у всіх PPTX)
    tblStyle.text = '{5940675A-B579-460E-94D1-54222C63F5DA}'

# --- ГЕНЕРАЦІЯ ДІАГРАМ ---
def create_chart_image(qs: QuestionSummary) -> io.BytesIO:
    plt.clf()
    plt.rcParams.update({'font.size': FONT_SIZE_CHART})
    
    labels = qs.table["Варіант відповіді"].astype(str).tolist()
    values = qs.table["Кількість"]
    wrapped_labels = [textwrap.fill(l, 25) for l in labels]

    if qs.question.qtype == QuestionType.SCALE:
        # Стовпчикова
        fig = plt.figure(figsize=(6.0, 4.5))
        bars = plt.bar(wrapped_labels, values, color='#4F81BD', width=BAR_WIDTH)
        plt.ylabel('Кількість')
        plt.grid(axis='y', linestyle='--', alpha=0.5)
        plt.xticks(rotation=0)
        
        for bar in bars:
            height = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                     f'{int(height)}', ha='center', va='bottom', fontweight='bold')
    else:
        # Кругова
        fig = plt.figure(figsize=(6.0, 5.0))
        colors = ['#4F81BD', '#C0504D', '#9BBB59', '#8064A2', '#4BACC6', '#F79646']
        c_arg = colors[:len(values)] if len(values) <= len(colors) else None
        
        wedges, texts, autotexts = plt.pie(
            values, labels=None, autopct='%1.1f%%', startangle=90,
            pctdistance=0.8, colors=c_arg, radius=1.1,
            textprops={'fontsize': FONT_SIZE_CHART}
        )
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_weight('bold')
            import matplotlib.patheffects as path_effects
            autotext.set_path_effects([path_effects.withStroke(linewidth=2, foreground='#333333')])

        plt.axis('equal')
        cols = 2 if len(labels) > 2 else 1
        plt.legend(wrapped_labels, loc="upper center", bbox_to_anchor=(0.5, 0.0), ncol=cols, frameon=False, fontsize=10)

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
    background_image_path: Optional[str] = "background.png"
) -> bytes:
    
    prs = Presentation()

    # --- ВСТАНОВЛЕННЯ ФОНУ ---
    # Застосовуємо фон до всіх доступних макетів (Layouts)
    if background_image_path and os.path.exists(background_image_path):
        for master in prs.slide_masters:
            # Фон майстра
            try:
                master.background.fill.user_picture(background_image_path)
            except: pass
            
            # Фон кожного лейауту в майстрі (це критично!)
            for layout in master.slide_layouts:
                try:
                    layout.background.fill.user_picture(background_image_path)
                except: pass

    # 1. Титульний
    try: slide_layout = prs.slide_layouts[0]
    except: slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    try:
        slide.shapes.title.text = "Звіт про результати опитування"
        slide.placeholders[1].text = f"Всього анкет: {len(original_df)}\nОброблено: {len(sliced_df)}\n{range_info}"
    except: pass

    # 2. Технічний
    try:
        slide_layout = prs.slide_layouts[1] 
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        title.text = "Технічна інформація"
        body_shape = slide.shapes.placeholders[1]
        tf = body_shape.text_frame
        
        for txt in [f"Загальна кількість респондентів: {len(original_df)}",
                    f"Кількість анкет у звіті: {len(sliced_df)}",
                    f"Діапазон обробки: {range_info}"]:
            p = tf.add_paragraph()
            p.text = txt
            p.font.size = Pt(20)
            p.level = 0
    except: pass

    # 3. Питання
    layout_index = 5 
    if len(prs.slide_layouts) <= 5: layout_index = len(prs.slide_layouts) - 1
    
    for qs in summaries:
        if qs.table.empty: continue
        
        slide = prs.slides.add_slide(prs.slide_layouts[layout_index])
        
        try:
            title = slide.shapes.title
            title.text = f"{qs.question.code}. {qs.question.text}"
            if len(title.text) > 60:
                title.text_frame.paragraphs[0].font.size = Pt(24)
        except: pass

        # --- ТАБЛИЦЯ ---
        rows = len(qs.table) + 1
        cols = 3
        table = slide.shapes.add_table(rows, cols, Inches(0.5), Inches(2.0), Inches(4.5), Inches(0.8)).table

        # ! ВАЖЛИВО: Застосовуємо стиль "Table Grid" (чорні рамки)
        apply_table_grid_style(table)

        # Ширина
        table.columns[0].width = Inches(2.5)
        table.columns[1].width = Inches(1.0)
        table.columns[2].width = Inches(1.0)

        # ХЕДЕР
        headers = ["Варіант", "Кільк.", "%"]
        for i, h in enumerate(headers):
            cell = table.cell(0, i)
            cell.text = h
            cell.text_frame.paragraphs[0].font.bold = True
            cell.text_frame.paragraphs[0].font.size = Pt(FONT_SIZE_TABLE_HEADER)
            
            # Фарбуємо заголовок (Світло-сірий)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(220, 220, 220)
            cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 0, 0)

        # ДАНІ
        for i, row in enumerate(qs.table.itertuples(index=False)):
            for j, val in enumerate(row):
                cell = table.cell(i+1, j)
                cell.text = str(val)
                cell.text_frame.paragraphs[0].font.size = Pt(FONT_SIZE_TABLE_DATA)
                
                # Вирівнювання чисел по центру
                if j > 0:
                    cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
                
                # Фарбуємо дані (Білий)
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(255, 255, 255)
                cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 0, 0)

        # --- ДІАГРАМА ---
        try:
            img_stream = create_chart_image(qs)
            slide.shapes.add_picture(img_stream, Inches(5.2), Inches(2.0), width=Inches(4.6))
        except: pass

    output = io.BytesIO()
    prs.save(output)
    return output.getvalue()