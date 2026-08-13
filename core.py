import io
import openpyxl
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

LEVELS = ["初級", "中級", "高級", "研經"]
FONT_NAME = "微軟正黑體"


def detect_level(filename: str):
    for lv in LEVELS:
        if lv in filename:
            return lv
    return None


def load_course_lookup(file_bytes, filename="course.xlsx"):
    """Returns {level_sheet_name: {mm/dd: course_name}}"""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    lookup = {}
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        date_to_course = {}
        for row in ws.iter_rows(min_row=3, values_only=False):
            date_cell = row[0].value
            course_cell = row[2].value if len(row) > 2 else None
            if date_cell is not None and hasattr(date_cell, "strftime"):
                mmdd = date_cell.strftime("%m/%d")
                date_to_course[mmdd] = course_cell if course_cell else "(課程名稱未提供)"
        lookup[sheet_name] = date_to_course
    return lookup


KNOWN_CODES = {"V", "O", "L", "LL", "M", "ML"}
ABSENCE_CODE = "O"


def _classify_status_cell(raw_value):
    """
    Classifies one attendance-status cell.
    Returns (kind, detail):
      kind = "absence"     -> counts toward makeup notice
      kind = "present"     -> recognized code, not an absence
      kind = "ambiguous"   -> multiple codes combined in one cell (e.g. "O / V")
      kind = "unrecognized"-> non-empty but not a known code
      kind = "empty"       -> no value recorded
    """
    if raw_value is None:
        return "empty", None
    s = str(raw_value).strip()
    if not s:
        return "empty", None

    if "/" in s:
        parts = [p.strip() for p in s.split("/") if p.strip()]
        if len(parts) > 1 and all(p in KNOWN_CODES for p in parts):
            return "ambiguous", s

    # allow a leading label like "夜研 V" — the code is the last whitespace token
    tokens = s.split()
    code = tokens[-1] if tokens else s

    if code == ABSENCE_CODE:
        return "absence", s
    if code in KNOWN_CODES:
        return "present", s
    return "unrecognized", s


def load_attendance(file_bytes, filename):
    """Returns (class_title, level, [student dicts], warnings) for one attendance file."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    title = ws["A1"].value or filename
    level = detect_level(filename)

    # date headers live in row 3, columns E..N (index 4..13, 0-based)
    header_row = next(ws.iter_rows(min_row=3, max_row=3, values_only=True))
    raw_date_headers = header_row[4:14]
    date_headers = []
    for h in raw_date_headers:
        if hasattr(h, "strftime"):
            date_headers.append(h.strftime("%m/%d"))
        elif h is None:
            date_headers.append("?")
        else:
            date_headers.append(str(h).strip())

    students = []
    warnings = []
    for row in ws.iter_rows(min_row=4, values_only=False):
        if len(row) < 2 or row[1].value is None:
            continue
        no_cell, name_cell, dharma_cell, group_cell = row[0], row[1], row[2], row[3]
        status_cells = row[4:14]

        absence_dates = []
        for col_idx, cell in enumerate(status_cells):
            mmdd = date_headers[col_idx] if col_idx < len(date_headers) else "?"
            kind, detail = _classify_status_cell(cell.value)
            if kind == "absence":
                absence_dates.append(mmdd)
            elif kind == "ambiguous":
                warnings.append(
                    f"{name_cell.value}（{filename}）於 {mmdd} 的簽到狀態為「{detail}」，"
                    f"格式模糊（同一格出現多種代碼），請人工確認是否算缺曠。"
                )
            elif kind == "unrecognized":
                warnings.append(
                    f"{name_cell.value}（{filename}）於 {mmdd} 的簽到狀態為「{detail}」，"
                    f"不是可辨識的代碼（V/O/L/LL/M/ML），請人工確認。"
                )

        if not absence_dates:
            continue

        students.append({
            "no": no_cell.value,
            "name": name_cell.value,
            "dharma_name": dharma_cell.value,
            "group": group_cell.value,
            "absence_count": len(absence_dates),
            "absence_dates": absence_dates,
            "level": level,
            "class_title": title,
            "source_file": filename,
        })
    return title, level, students, warnings


def build_student_records(attendance_files, course_lookup):
    """
    attendance_files: list of (filename, bytes)
    course_lookup: dict from load_course_lookup
    Returns list of student dicts, each with resolved 'items': [{date, course}]
    """
    all_students = []
    warnings = []
    for filename, file_bytes in attendance_files:
        title, level, students, cell_warnings = load_attendance(file_bytes, filename)
        warnings.extend(cell_warnings)
        date_map = course_lookup.get(level, {})
        if level is None:
            warnings.append(f"檔名「{filename}」無法判斷班級（初級/中級/高級/研經），將無法對應課程名稱。")
        for s in students:
            items = []
            for d in s["absence_dates"]:
                course = date_map.get(d)
                if course is None:
                    course = "(查無對應課程，請人工確認)"
                    warnings.append(f"{s['name']}（{filename}）的缺曠日期 {d} 在補課名單中找不到對應課程。")
                items.append({"date": d, "course": course})
            s["items"] = items
        all_students.extend(students)
    return all_students, warnings


# ---------------- docx generation (python-docx) ----------------

def _set_table_borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "999999")
        borders.append(el)
    tblPr.append(borders)


def _set_cell_shading(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    tcPr.append(shd)


def _style_run(run, size=11, bold=False):
    run.font.name = FONT_NAME
    run.font.size = Pt(size)
    run.font.bold = bold
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), FONT_NAME)


def _add_paragraph(doc, text="", size=11, bold=False, align=None, space_after=6):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(text)
    _style_run(run, size=size, bold=bold)
    return p


TEMPLATE_FONT = "標楷體"
TITLE_SIZE = 18
BODY_SIZE = 14
ORG_NAME = "普印精舍"

DEFAULT_DEADLINE = "9月2日"
DEFAULT_HOURS_NOTE = "於周一至周日早上9:00~20:00至精舍櫃檯報到完成補課，"
DEFAULT_LATE_NOTE = "若超過補課期限，請事先預約補課時間，以利事前安排。"


def _tmpl_run(paragraph, text, size=BODY_SIZE, bold=True, underline=False):
    run = paragraph.add_run(text)
    run.font.name = TEMPLATE_FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.underline = underline
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), TEMPLATE_FONT)
    return run


def _add_text_with_underlined_phrases(paragraph, text, underline_phrases):
    """
    Writes `text` into `paragraph`, automatically underlining the first
    occurrence of any phrase in `underline_phrases` that appears in it —
    so the original template's underline styling (e.g. under "事先預約")
    survives even if the surrounding wording is edited.
    """
    remaining = text
    offset = 0
    match_start, match_phrase = None, None
    for phrase in underline_phrases:
        if phrase and phrase in remaining:
            i = remaining.index(phrase)
            if match_start is None or i < match_start:
                match_start, match_phrase = i, phrase
    if match_phrase is None:
        _tmpl_run(paragraph, text)
        return
    before = remaining[:match_start]
    after = remaining[match_start + len(match_phrase):]
    if before:
        _tmpl_run(paragraph, before)
    _tmpl_run(paragraph, match_phrase, underline=True)
    if after:
        _tmpl_run(paragraph, after)


def build_notice_docx(student, org_name=ORG_NAME, deadline_text=DEFAULT_DEADLINE,
                       hours_note=DEFAULT_HOURS_NOTE, late_note=DEFAULT_LATE_NOTE):
    """
    Build a single-student makeup-class notice matching the organization's
    official Word template (普印精舍-補課通知單).
    Blanks in the original hand-filled template are replaced with the
    student's actual data; the number of 缺課日期/缺課名稱 lines matches
    that student's actual absence count.
    """
    doc = Document()
    section = doc.sections[0]
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)
    section.top_margin = Cm(2.5)

    # Title
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _tmpl_run(p, f"{org_name}–補課通知單", size=TITLE_SIZE)

    doc.add_paragraph()

    # 班別/姓名/法名
    p = doc.add_paragraph()
    _tmpl_run(
        p,
        f"班別：{student.get('level') or ''}　　姓名：{student.get('name','')}　　"
        f"法名：{student.get('dharma_name') or ''}",
    )

    doc.add_paragraph()

    # 缺課日期／缺課名稱 lines — one per actual absence
    for item in student.get("items", []):
        p = doc.add_paragraph()
        _tmpl_run(p, f"缺課日期：{item['date']}　　缺課名稱：{item['course']}")

    doc.add_paragraph()

    # Deadline notice (deadline date underlined, matching original template)
    p = doc.add_paragraph()
    _tmpl_run(p, "請在")
    _tmpl_run(p, deadline_text, underline=True)
    _tmpl_run(p, "前，")
    _add_text_with_underlined_phrases(p, hours_note, underline_phrases=[])

    p = doc.add_paragraph()
    _add_text_with_underlined_phrases(p, late_note, underline_phrases=["事先預約"])

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf
