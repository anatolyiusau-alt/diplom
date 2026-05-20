#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генератор Word-документу бакалаврської роботи (ЗНУ, кафедра політології, 2026).

Дотримується:
  - Методичних вказівок ЗНУ 2025 (Вагіна, Горло, Мальована);
  - ДСТУ 3008:2015 («Звіти у сфері науки і техніки»);
  - Правок наукового керівника до плану.

Параметри оформлення:
  - Поля: ліве 30 мм, праве 10 мм, верхнє 20 мм, нижнє 20 мм
  - Шрифт: Times New Roman 14
  - Інтервал: 1.5
  - Абзацний відступ: 1.25 см
  - Вирівнювання тексту: за шириною
  - Заголовки розділів (ЗМІСТ, ВСТУП, РОЗДІЛ N, ВИСНОВКИ, СПИСОК ЛІТЕРАТУРИ,
    ДОДАТКИ): великі літери, центр, 14 напівжирний, без крапки
  - Підрозділи (1.1., 2.3., …): з абзацного відступу, 14 напівжирний
  - Лапки: « » (auto-conversion)
  - Тире: — між словами; -, дефіс — у складених словах (без перетворення)
  - Курсив для термінів вступу: об'єкт, предмет, мета, гіпотеза, методи
  - Кожен новий розділ — з нової сторінки
  - Нумерація сторінок: правий верхній кут, арабські цифри без крапки

Запуск:
    python3 build/build.py
    → output/diplom.docx
"""

from __future__ import annotations

import os
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm, Pt, RGBColor

# Дозволяємо імпорт config.py з тієї ж директорії
sys.path.insert(0, str(Path(__file__).resolve().parent))
import config as cfg  # noqa: E402

# ──────────────────────────────────────────────────────────────────────────────
# КОНСТАНТИ ОФОРМЛЕННЯ
# ──────────────────────────────────────────────────────────────────────────────

FONT_NAME = "Times New Roman"
FONT_SIZE = Pt(14)
LINE_SPACING = 1.5
FIRST_LINE_INDENT = Cm(1.25)

MARGIN_LEFT = Mm(30)
MARGIN_RIGHT = Mm(10)
MARGIN_TOP = Mm(20)
MARGIN_BOTTOM = Mm(20)

# Літери для додатків (без Ґ, Є, З, І, Ї, Й, О, Ч, Ь — за п. 3.7 методички)
APPENDIX_LETTERS = list("АБВГДЕЖИКЛМНПРСТУФХЦШЩЮЯ")


# ──────────────────────────────────────────────────────────────────────────────
# НИЗЬКОРІВНЕВІ ХЕЛПЕРИ (XML-маніпуляції)
# ──────────────────────────────────────────────────────────────────────────────

def set_run_font(run, *, bold=False, italic=False, size=FONT_SIZE, upper=False):
    """Встановити шрифт TNR із заданими параметрами."""
    run.font.name = FONT_NAME
    # Для кирилиці треба явно задати шрифт через rPr/eastAsia/cs
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), FONT_NAME)
    rFonts.set(qn("w:hAnsi"), FONT_NAME)
    rFonts.set(qn("w:cs"), FONT_NAME)
    rFonts.set(qn("w:eastAsia"), FONT_NAME)
    run.font.size = size
    run.bold = bold
    run.italic = italic
    if upper:
        run.font.all_caps = True


def add_page_number_field(paragraph):
    """Додати поле PAGE у параграф (для нумерації сторінок)."""
    fldChar1 = OxmlElement("w:fldChar")
    fldChar1.set(qn("w:fldCharType"), "begin")

    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = "PAGE   \\* MERGEFORMAT"

    fldChar2 = OxmlElement("w:fldChar")
    fldChar2.set(qn("w:fldCharType"), "end")

    run = paragraph.add_run()
    set_run_font(run)
    run._element.append(fldChar1)
    run._element.append(instrText)
    run._element.append(fldChar2)


def add_page_break(doc):
    """Додати розрив сторінки."""
    p = doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)


def add_section_break(doc, *, start_type="nextPage"):
    """Додати розрив секції (для незалежних колонтитулів)."""
    new_section = doc.add_section(WD_SECTION.NEW_PAGE)
    return new_section


def configure_section_margins(section):
    """Виставити поля 30/10/20/20 мм для секції."""
    section.left_margin = MARGIN_LEFT
    section.right_margin = MARGIN_RIGHT
    section.top_margin = MARGIN_TOP
    section.bottom_margin = MARGIN_BOTTOM
    section.page_width = Mm(210)   # A4
    section.page_height = Mm(297)  # A4


def set_page_number_in_header(section, *, hidden=False):
    """Додати номер сторінки у правий верхній кут.

    Якщо hidden=True — поле є, але прозоре (для першої сторінки секції).
    """
    header = section.header
    header.is_linked_to_previous = False
    p = header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.first_line_indent = Cm(0)
    if hidden:
        return  # просто не додаємо поле — буде порожній header
    add_page_number_field(p)


# ──────────────────────────────────────────────────────────────────────────────
# СТИЛІ ПАРАГРАФІВ
# ──────────────────────────────────────────────────────────────────────────────

def style_normal_paragraph(p):
    """Звичайний абзац: TNR 14, 1.5 інтервал, 1.25 абзац, justified."""
    pf = p.paragraph_format
    pf.first_line_indent = FIRST_LINE_INDENT
    pf.line_spacing = LINE_SPACING
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


def style_centered_paragraph(p, *, indent=False):
    """Центрований абзац (для титулки, заголовків)."""
    pf = p.paragraph_format
    pf.first_line_indent = FIRST_LINE_INDENT if indent else Cm(0)
    pf.line_spacing = LINE_SPACING
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER


# ──────────────────────────────────────────────────────────────────────────────
# ВИСОКОРІВНЕВЕ API: додавання структурних елементів
# ──────────────────────────────────────────────────────────────────────────────

def add_section_heading(doc, text):
    """Заголовок розділу: ЗМІСТ, ВСТУП, РОЗДІЛ N, ВИСНОВКИ, ...

    Великі літери, центр, 14 напівжирний.
    """
    p = doc.add_paragraph()
    style_centered_paragraph(p)
    run = p.add_run(text.upper())
    set_run_font(run, bold=True)
    return p


def add_subsection_heading(doc, text):
    """Заголовок підрозділу: «1.1. Назва...»

    З абзацного відступу, 14 напівжирний, без крапки в кінці.
    """
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.first_line_indent = FIRST_LINE_INDENT
    pf.line_spacing = LINE_SPACING
    pf.space_before = Pt(12)
    pf.space_after = Pt(6)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    text = text.rstrip(".").strip()
    run = p.add_run(text)
    set_run_font(run, bold=True)
    return p


def add_normal_paragraph(doc, text, *, italic_terms=None):
    """Звичайний абзац з опц. виділенням курсивом окремих слів.

    italic_terms: список слів/фраз, які треба виділити курсивом
                  (для термінів у вступі: об'єкт, предмет, мета, ...).
    """
    p = doc.add_paragraph()
    style_normal_paragraph(p)
    italic_terms = italic_terms or []
    if italic_terms:
        # Розбиваємо текст по курсивних термінах
        pattern = "|".join(re.escape(t) for t in italic_terms)
        parts = re.split(f"({pattern})", text)
        for part in parts:
            run = p.add_run(part)
            is_italic = part in italic_terms
            set_run_font(run, italic=is_italic)
    else:
        # Може містити inline-виділення *italic* і **bold**
        for run_text, fmt in _parse_inline(text):
            run = p.add_run(run_text)
            set_run_font(run, bold=fmt.get("bold", False), italic=fmt.get("italic", False))
    return p


def add_centered_text(doc, text, *, bold=False, upper=False):
    """Центрований рядок (для титулки)."""
    p = doc.add_paragraph()
    style_centered_paragraph(p)
    run = p.add_run(text.upper() if upper else text)
    set_run_font(run, bold=bold)
    return p


def add_blank_line(doc):
    """Порожній рядок-розділювач."""
    p = doc.add_paragraph()
    style_normal_paragraph(p)
    return p


# ──────────────────────────────────────────────────────────────────────────────
# ПАРСЕР MARKDOWN (мінімальний)
# ──────────────────────────────────────────────────────────────────────────────

INLINE_BOLD_RE = re.compile(r"\*\*(.+?)\*\*")
INLINE_ITALIC_RE = re.compile(r"(?<!\*)\*([^*]+?)\*(?!\*)")


def _parse_inline(text):
    """Розпарсити inline **bold** і *italic*. Повертає список (text, format)."""
    # Спочатку шукаємо **bold**, потім всередині — *italic*
    tokens = []
    pos = 0
    pattern = re.compile(r"(\*\*([^*]+?)\*\*|(?<!\*)\*([^*]+?)\*(?!\*))")
    for m in pattern.finditer(text):
        if m.start() > pos:
            tokens.append((text[pos:m.start()], {}))
        if m.group(2) is not None:  # bold
            tokens.append((m.group(2), {"bold": True}))
        elif m.group(3) is not None:  # italic
            tokens.append((m.group(3), {"italic": True}))
        pos = m.end()
    if pos < len(text):
        tokens.append((text[pos:], {}))
    return tokens or [(text, {})]


def normalize_typography(text):
    """Українська типографіка:
    - "..."  →  «...»
    - --     →  —
    - ' '    →  обережно (не чіпаємо англ. апострофи)
    Робиться рядок за рядком до парсингу.
    """
    # Тире
    text = re.sub(r"(?<=\s)--(?=\s)", "—", text)
    # Лапки: попарно перетворюємо " ... " на « ... »
    # Простий алгоритм: чергуємо відкривальну і закривальну
    out = []
    open_q = True
    for ch in text:
        if ch == '"':
            out.append("«" if open_q else "»")
            open_q = not open_q
        else:
            out.append(ch)
    return "".join(out)


def render_markdown_to_docx(doc, md_text, *, italic_terms=None):
    """Перетворити простий markdown на параграфи у docx.

    Підтримуються:
      # SECTION HEADING        →  ЗАГОЛОВОК РОЗДІЛУ (ВСТУП, РОЗДІЛ N, ...)
      ## 1.1. Subheading       →  заголовок підрозділу
      ### Inline emphasis      →  жирний абзац (для дрібних виділень)
      звичайний текст          →  абзац (із підтримкою *italic*, **bold**)
      <pagebreak/>             →  розрив сторінки
      <center>...</center>     →  центрований абзац
      <center bold>...</...>   →  центрований напівжирний
    """
    md_text = normalize_typography(md_text)
    lines = md_text.split("\n")
    buf = []  # буфер для абзацу

    def flush_paragraph():
        if not buf:
            return
        joined = " ".join(s.strip() for s in buf if s.strip())
        if joined:
            add_normal_paragraph(doc, joined, italic_terms=italic_terms)
        buf.clear()

    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if not stripped:
            flush_paragraph()
            i += 1
            continue

        # <pagebreak/>
        if stripped == "<pagebreak/>":
            flush_paragraph()
            add_page_break(doc)
            i += 1
            continue

        # <center>...</center> або <center bold>...</center>
        m = re.match(r"<center( bold)?( upper)?>(.*?)</center>", stripped)
        if m:
            flush_paragraph()
            bold = bool(m.group(1))
            upper = bool(m.group(2))
            add_centered_text(doc, m.group(3), bold=bold, upper=upper)
            i += 1
            continue

        # # Section heading
        if stripped.startswith("# ") and not stripped.startswith("## "):
            flush_paragraph()
            add_section_heading(doc, stripped[2:].strip())
            i += 1
            continue

        # ## Subheading
        if stripped.startswith("## ") and not stripped.startswith("### "):
            flush_paragraph()
            add_subsection_heading(doc, stripped[3:].strip())
            i += 1
            continue

        # ### Bold abzac
        if stripped.startswith("### "):
            flush_paragraph()
            p = doc.add_paragraph()
            style_normal_paragraph(p)
            run = p.add_run(stripped[4:].strip())
            set_run_font(run, bold=True)
            i += 1
            continue

        # звичайний рядок
        buf.append(line)
        i += 1

    flush_paragraph()


# ──────────────────────────────────────────────────────────────────────────────
# ШАБЛОНИ СТРУКТУРНИХ ЕЛЕМЕНТІВ ЗА МЕТОДИЧКОЮ
# ──────────────────────────────────────────────────────────────────────────────

def render_titulka(doc):
    """Титульний аркуш (за зразком Додатку А методички)."""
    section = doc.sections[0]
    configure_section_margins(section)

    add_centered_text(doc, cfg.MINISTRY)
    add_centered_text(doc, cfg.UNIVERSITY, bold=True)
    add_centered_text(doc, cfg.FACULTY)
    add_centered_text(doc, cfg.DEPARTMENT)

    for _ in range(8):
        add_blank_line(doc)

    add_centered_text(doc, "Кваліфікаційна робота", bold=True)
    add_centered_text(doc, cfg.DEGREE_LEVEL)
    add_centered_text(doc, "на тему:")
    add_centered_text(doc, f"«{cfg.THESIS_TITLE.upper()}»", bold=True)

    for _ in range(4):
        add_blank_line(doc)

    # Виконавець / керівник — лівий блок
    block = [
        f"Виконав: студент {cfg.STUDENT_COURSE} курсу,",
        f"групи {cfg.STUDENT_GROUP},",
        f"спеціальності {cfg.SPECIALTY_CODE} «{cfg.SPECIALTY_NAME}»,",
        f"освітньо-професійної програми",
        f"«{cfg.EDUCATIONAL_PROGRAM}»",
        f"{cfg.STUDENT_INITIALS}",
        "",
        f"Керівник: {cfg.SUPERVISOR_RANK}",
        f"{cfg.SUPERVISOR_INITIALS}",
        "",
        f"Рецензент: {cfg.REVIEWER_RANK}",
        f"{cfg.REVIEWER_INITIALS}",
    ]
    for line in block:
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.first_line_indent = Cm(0)
        pf.left_indent = Cm(8)  # зсув праворуч, типово для титулки
        pf.line_spacing = 1.0
        pf.space_after = Pt(0)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(line)
        set_run_font(run)

    # Внизу — місто і рік
    for _ in range(2):
        add_blank_line(doc)
    add_centered_text(doc, f"{cfg.CITY} – {cfg.DEFENSE_YEAR}")


def render_declaration(doc):
    """Декларація академічної доброчесності."""
    add_centered_text(doc, "ДЕКЛАРАЦІЯ", bold=True, upper=True)
    add_centered_text(doc, "АКАДЕМІЧНОЇ ДОБРОЧЕСНОСТІ", bold=True, upper=True)
    add_centered_text(doc, "ЗДОБУВАЧА СТУПЕНЯ ВИЩОЇ ОСВІТИ ЗНУ", bold=True, upper=True)
    add_blank_line(doc)

    text = (
        f"Я, {cfg.STUDENT_FULL_NAME}, студент {cfg.STUDENT_COURSE} курсу, форми навчання денної, "
        f"{cfg.FACULTY.lower()}, спеціальності {cfg.SPECIALTY_CODE} «{cfg.SPECIALTY_NAME}», "
        f"освітньо-професійної програми «{cfg.EDUCATIONAL_PROGRAM}», "
        f"адреса електронної пошти {cfg.STUDENT_EMAIL or '___________________'}, — підтверджую, що написана мною "
        f"кваліфікаційна робота на тему «{cfg.THESIS_TITLE}» відповідає вимогам академічної доброчесності "
        f"та не містить порушень, що визначені у ст. 42 Закону України «Про освіту», "
        f"зі змістом яких ознайомлений/ознайомлена."
    )
    add_normal_paragraph(doc, text)
    add_blank_line(doc)
    add_normal_paragraph(
        doc,
        "Заявляю, що надана мною для перевірки електронна версія роботи є ідентичною її друкованій версії.",
    )
    add_blank_line(doc)
    add_normal_paragraph(
        doc,
        "Згоден/згодна на перевірку моєї роботи на відповідність критеріям академічної доброчесності "
        "у будь-який спосіб, у тому числі за допомогою інтернет-системи, а також на архівування "
        "моєї роботи в базі даних цієї системи.",
    )

    for _ in range(3):
        add_blank_line(doc)

    # Підпис
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0)
    run = p.add_run("Дата: __________            Підпис: __________            ")
    set_run_font(run)
    run = p.add_run(cfg.STUDENT_INITIALS)
    set_run_font(run)


def render_zmist(doc):
    """Зміст (генерується вручну — у Word краще оновити поле TOC)."""
    add_section_heading(doc, "ЗМІСТ")
    add_blank_line(doc)

    items = [
        ("ВСТУП", "—"),
        ("РОЗДІЛ 1. ТЕОРЕТИКО-МЕТОДОЛОГІЧНІ ЗАСАДИ ДОСЛІДЖЕННЯ СОЦІАЛЬНИХ МЕРЕЖ ЯК ІНСТРУМЕНТУ ФОРМУВАННЯ ГРОМАДСЬКОЇ ДУМКИ", "—"),
        ("    1.1. Проблематика дослідження соціальних мереж та громадської думки у сучасній політичній науці", "—"),
        ("    1.2. Аналіз основних понять дослідження", "—"),
        ("    1.3. Методологія дослідження соціальних мереж як інструменту впливу на громадську думку", "—"),
        ("РОЗДІЛ 2. ТЕОРЕТИЧНІ АСПЕКТИ ДОСЛІДЖЕННЯ ВПЛИВУ СОЦІАЛЬНИХ МЕРЕЖ НА ГРОМАДСЬКУ ДУМКУ В УМОВАХ ВІЙНИ", "—"),
        ("    2.1. Основні платформи соціальних мереж та їхній комунікаційний потенціал у воєнний період", "—"),
        ("    2.2. Ключові суб'єкти формування громадської думки в соціальних мережах", "—"),
        ("    2.3. Конструктивний та деструктивний вплив соціальних мереж на громадську думку в умовах війни", "—"),
        ("РОЗДІЛ 3. ОСОБЛИВОСТІ ВПЛИВУ СОЦІАЛЬНИХ МЕРЕЖ НА ФОРМУВАННЯ ГРОМАДСЬКОЇ ДУМКИ В УКРАЇНІ В УМОВАХ ВІЙНИ", "—"),
        ("    3.1. Інформаційний простір України в умовах війни як середовище формування громадської думки в соціальних мережах", "—"),
        ("    3.2. Ключові наративи, інструменти та практики впливу в українському сегменті соціальних мереж", "—"),
        ("    3.3. Шляхи підвищення ефективності та безпечності використання соціальних мереж у формуванні громадської думки в Україні", "—"),
        ("ВИСНОВКИ", "—"),
        ("СПИСОК ЛІТЕРАТУРИ", "—"),
        ("ДОДАТКИ", "—"),
    ]
    for title, page in items:
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.first_line_indent = Cm(0)
        pf.line_spacing = LINE_SPACING
        pf.space_after = Pt(0)
        # Tab з лідером — крапки до правого краю
        from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER
        pf.tab_stops.add_tab_stop(Cm(16), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        run = p.add_run(title)
        set_run_font(run)
        run = p.add_run("\t" + page)
        set_run_font(run)


# ──────────────────────────────────────────────────────────────────────────────
# ЗБІР ВМІСТУ З content/
# ──────────────────────────────────────────────────────────────────────────────

CONTENT_DIR = Path(__file__).resolve().parent.parent / "content"
OUTPUT_DIR = Path(__file__).resolve().parent.parent / "output"

# Терміни вступу, які треба виділяти курсивом
VSTUP_ITALIC_TERMS = [
    "Об'єкт дослідження",
    "Предмет дослідження",
    "Мета дослідження",
    "Гіпотеза дослідження",
    "Завдання дослідження",
    "Методи дослідження",
    "об'єктом дослідження",
    "предметом дослідження",
    "метою дослідження",
    "гіпотезою дослідження",
    "методами дослідження",
]


def load_content_file(name):
    """Прочитати content/<name>.md → str. Якщо файла немає — повернути плейсхолдер."""
    path = CONTENT_DIR / f"{name}.md"
    if not path.exists():
        return f"_(розділ «{name}» поки не написано — додайте файл content/{name}.md)_\n"
    return path.read_text(encoding="utf-8")


# ──────────────────────────────────────────────────────────────────────────────
# ГОЛОВНА ЗБІРКА ДОКУМЕНТА
# ──────────────────────────────────────────────────────────────────────────────

def build_document():
    doc = Document()

    # Налаштувати стиль Normal глобально
    style = doc.styles["Normal"]
    style.font.name = FONT_NAME
    style.font.size = FONT_SIZE
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), FONT_NAME)
    rFonts.set(qn("w:hAnsi"), FONT_NAME)
    rFonts.set(qn("w:cs"), FONT_NAME)
    rFonts.set(qn("w:eastAsia"), FONT_NAME)

    # Поля для першої секції
    configure_section_margins(doc.sections[0])

    # Нумерація сторінок (правий верхній кут)
    set_page_number_in_header(doc.sections[0])

    # 1. Титулка
    render_titulka(doc)
    add_page_break(doc)

    # 2. Декларація академічної доброчесності
    render_declaration(doc)
    add_page_break(doc)

    # 3. Технічне завдання
    tz_md = load_content_file("02_tz")
    render_markdown_to_docx(doc, tz_md)
    add_page_break(doc)

    # 4. Реферат українською
    referat_md = load_content_file("03_referat_ua")
    render_markdown_to_docx(doc, referat_md)
    add_page_break(doc)

    # 5. Abstract англійською
    abstract_md = load_content_file("04_abstract_en")
    render_markdown_to_docx(doc, abstract_md)
    add_page_break(doc)

    # 6. Зміст
    render_zmist(doc)
    add_page_break(doc)

    # 7. Вступ
    vstup_md = load_content_file("06_vstup")
    render_markdown_to_docx(doc, vstup_md, italic_terms=VSTUP_ITALIC_TERMS)
    add_page_break(doc)

    # 8. Розділ 1
    rozdil_1_md = load_content_file("07_rozdil_1")
    render_markdown_to_docx(doc, rozdil_1_md)
    add_page_break(doc)

    # 9. Розділ 2
    rozdil_2_md = load_content_file("08_rozdil_2")
    render_markdown_to_docx(doc, rozdil_2_md)
    add_page_break(doc)

    # 10. Розділ 3
    rozdil_3_md = load_content_file("09_rozdil_3")
    render_markdown_to_docx(doc, rozdil_3_md)
    add_page_break(doc)

    # 11. Висновки
    vysnovky_md = load_content_file("10_vysnovky")
    render_markdown_to_docx(doc, vysnovky_md)
    add_page_break(doc)

    # 12. Список літератури
    literatura_md = load_content_file("11_literatura")
    render_markdown_to_docx(doc, literatura_md)
    add_page_break(doc)

    # 13. Додатки
    dodatky_md = load_content_file("12_dodatky")
    render_markdown_to_docx(doc, dodatky_md)

    return doc


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    doc = build_document()
    out_path = OUTPUT_DIR / "diplom.docx"
    doc.save(out_path)
    print(f"OK  {out_path}  ({out_path.stat().st_size / 1024:.1f} KB)")


if __name__ == "__main__":
    main()
