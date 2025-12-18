import io
import textwrap
import pandas as pd
import matplotlib
matplotlib.use('Agg') 
import matplotlib.pyplot as plt

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn

from classification import QuestionInfo, QuestionType
from summary import QuestionSummary
from typing import List, Optional

# --- НАЛАШТУВАННЯ ---
CHART_DPI = 150
FONT_SIZE_CHART = 11        
FONT_SIZE_HEADER = 12 
FONT_SIZE_DATA = 11   
BAR_WIDTH = 0.6

def set_table_grid_style(table):
    """Вмикає чорні рамки (Table Grid)."""
    tbl = table._tbl
    tblPr = tbl.tblPr
    tblStyle = tblPr.find(qn('a:tableStyleId'))
    if tblStyle is None:
        tblStyle = OxmlElement('a:tableStyleId')
        tblPr.append(tblStyle)
    tblStyle.text = '{5940675A-B579-460E-94D1-54222C63F5DA}'

def create_chart_image(qs: QuestionSummary) -> io.BytesIO:
    """
    Генерує зображення діаграми з жорстким очищенням пам'яті.
    Це гарантує, що в ZIP-архіві діаграми не переплутаються.
    """
    # 1. Жорстке очищення попередніх графіків
    plt.close('all') 
    
    # 2. Налаштування шрифту
    plt.rcParams.update({'font.size': FONT_SIZE_CHART})
    
    labels = qs.table["Варіант відповіді"].astype(str).tolist()
    values = qs.table["Кількість"]
    wrapped_labels = [textwrap.fill(l, 25) for l in labels]

    # 3. Створення нової фігури
    if qs.question.qtype == QuestionType.SCALE:
        # --- СТОВПЧИКОВА (BAR) ---
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
        # --- КРУГОВА (PIE) ---
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
    
    # 4. Збереження
    img_stream = io.BytesIO()
    plt.savefig(img_stream, format='png', dpi=CHART_DPI, bbox_inches='tight')
    
    # 5. Фінальне очищення цієї конкретної фігури
    plt.close(fig) 
    img_stream.seek(0)
    return img_stream

def build_pptx_report(
    original_df: pd.DataFrame,
    sliced_df: pd.DataFrame,
    summaries: List[QuestionSummary],
    range_info: str
) -> bytes:
    
    prs = Presentation()

    # Слайд 1: Титул
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    try:
        slide.shapes.title.text = "Звіт про результати опитування"
        slide.placeholders[1].text = f"Всього анкет: {len(original_df)}\nОброблено: {len(sliced_df)}\n{range_info}"
    except: pass

    # Слайд 2: Технічний
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    try:
        slide.shapes.title.text = "Технічна інформація"
        tf = slide.placeholders[1].text_frame
        tf.text = "Параметри вибірки:"
        for item in [f"Загальна кількість: {len(original_df)}", f"У звіті: {len(sliced_df)}", f"Діапазон: {range_info}"]:
            p = tf.add_paragraph()
            p.text = item
            p.font.size = Pt(20)
    except: pass

    # Слайди даних
    layout_index = 5 
    if len(prs.slide_layouts) <= 5: layout_index = len(prs.slide_layouts) - 1
    
    for qs in summaries:
        if qs.table.empty: continue
        
        slide = prs.slides.add_slide(prs.slide_layouts[layout_index])
        
        try:
            title = slide.shapes.title
            title.text = f"{qs.question.code}. {qs.question.text}"
            title.text_frame.paragraphs[0].font.size = Pt(24) if len(title.text) > 60 else Pt(32)
        except: pass

        # Таблиця
        rows = len(qs.table) + 1
        cols = 3
        table = slide.shapes.add_table(rows, cols, Inches(0.5), Inches(2.0), Inches(4.5), Inches(0.8)).table
        set_table_grid_style(table) # Рамки

        table.columns[0].width = Inches(2.5)
        table.columns[1].width = Inches(1.0)
        table.columns[2].width = Inches(1.0)

        # Хедер
        headers = ["Варіант", "Кільк.", "%"]
        for i, h in enumerate(headers):
            cell = table.cell(0, i)
            cell.text = h
            cell.text_frame.paragraphs[0].font.bold = True
            cell.text_frame.paragraphs[0].font.size = Pt(FONT_SIZE_HEADER)
            cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(220, 220, 220) # Сірий
            cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 0, 0)

        # Дані
        for i, row in enumerate(qs.table.itertuples(index=False)):
            for j, val in enumerate(row):
                cell = table.cell(i+1, j)
                cell.text = str(val)
                cell.text_frame.paragraphs[0].font.size = Pt(FONT_SIZE_DATA)
                cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 0, 0)
                
                if j > 0: cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
                else: cell.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
                
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(255, 255, 255) # Білий

        # Діаграма
        try:
            img_stream = create_chart_image(qs)
            slide.shapes.add_picture(img_stream, Inches(5.2), Inches(2.0), width=Inches(4.6))
        except: pass

    output = io.BytesIO()
    prs.save(output)
    return output.getvalue()