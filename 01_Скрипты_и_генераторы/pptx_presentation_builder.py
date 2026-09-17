"""
Универсальный генератор PowerPoint (.pptx) презентаций для академических докладов.
Создает 100% редактируемые слайды с нативными блоками, списками и таблицами MS PowerPoint.
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE


class EditablePPTXBuilder:
    """Конструктор редактируемых слайдов 16:9 с нативными объектами PowerPoint."""

    def __init__(self, theme="dark"):
        self.prs = Presentation()
        # 16:9 Widescreen (13.333 x 7.5 inches)
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)
        self.blank_layout = self.prs.slide_layouts[6]

        if theme == "light":
            self.COLOR_BG = RGBColor(248, 250, 252)
            self.COLOR_CARD = RGBColor(255, 255, 255)
            self.COLOR_CARD_BORDER = RGBColor(226, 232, 240)
            self.COLOR_TEXT_MAIN = RGBColor(15, 23, 42)
            self.COLOR_TEXT_MUTED = RGBColor(71, 85, 105)
            self.COLOR_PRIMARY = RGBColor(2, 132, 199)
            self.COLOR_ACCENT = RGBColor(99, 102, 241)
        else:  # dark (default)
            self.COLOR_BG = RGBColor(11, 15, 25)
            self.COLOR_CARD = RGBColor(22, 31, 51)
            self.COLOR_CARD_BORDER = RGBColor(40, 55, 85)
            self.COLOR_TEXT_MAIN = RGBColor(248, 250, 252)
            self.COLOR_TEXT_MUTED = RGBColor(148, 163, 184)
            self.COLOR_PRIMARY = RGBColor(56, 189, 248)
            self.COLOR_ACCENT = RGBColor(129, 140, 248)

    def add_blank_slide(self):
        """Добавляет пустой слайд с заливкой фона."""
        slide = self.prs.slides.add_slide(self.blank_layout)
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, self.prs.slide_width, self.prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = self.COLOR_BG
        bg.line.fill.background()
        return slide

    def add_header(self, slide, tag_text, title_text, subtitle_text="", title_size=26):
        """Добавляет верхний блок слайда: тег раздела, заголовок и подзаголовок."""
        tag_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.7), Inches(0.4))
        tf = tag_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = tag_text.upper()
        p.font.name = "Arial"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = self.COLOR_PRIMARY

        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.85), Inches(11.7), Inches(0.8))
        tf = title_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title_text
        p.font.name = "Arial"
        p.font.size = Pt(title_size)
        p.font.bold = True
        p.font.color.rgb = self.COLOR_TEXT_MAIN

        if subtitle_text:
            sub_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.55), Inches(11.7), Inches(0.4))
            tf = sub_box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = subtitle_text
            p.font.name = "Arial"
            p.font.size = Pt(13)
            p.font.color.rgb = self.COLOR_TEXT_MUTED

    def add_footer(self, slide, slide_num, total_slides=12):
        """Добавляет нижний колонтитул с нумерацией слайдов."""
        footer_box = slide.shapes.add_textbox(Inches(0.8), Inches(6.9), Inches(11.7), Inches(0.4))
        tf = footer_box.text_frame
        p = tf.paragraphs[0]
        p.text = f"ФГАОУ ВО МГТУ «СТАНКИН»  |  Слайд {slide_num:02d} / {total_slides:02d}"
        p.font.name = "Arial"
        p.font.size = Pt(10)
        p.font.color.rgb = self.COLOR_TEXT_MUTED

    def add_card(self, slide, left, top, width, height, title, bullet_items, icon="", border_color=None, title_size=17):
        """Добавляет нативную карточку с заголовком и маркированным списком."""
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = self.COLOR_CARD
        card.line.color.rgb = border_color or self.COLOR_CARD_BORDER
        card.line.width = Pt(1.5)

        tb = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.15), width - Inches(0.4), height - Inches(0.3))
        tf = tb.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = f"{icon}  {title}" if icon else title
        p.font.name = "Arial"
        p.font.size = Pt(title_size)
        p.font.bold = True
        p.font.color.rgb = self.COLOR_TEXT_MAIN
        p.space_after = Pt(10)

        for item in bullet_items:
            p_item = tf.add_paragraph()
            p_item.text = f"• {item}"
            p_item.font.name = "Arial"
            p_item.font.size = Pt(12)
            p_item.font.color.rgb = self.COLOR_TEXT_MUTED
            p_item.space_after = Pt(6)

    def save(self, filepath):
        """Сохраняет презентацию, создавая при необходимости целевую директорию."""
        os.makedirs(os.path.dirname(os.path.abspath(filepath)), exist_ok=True)
        self.prs.save(filepath)
        print(f"[PPTX Builder] Презентация успешно сохранена: {filepath}")
