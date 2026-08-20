import argparse
import json
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "assets"


def insert_before_any(parent, element, local_names):
    tags = {qn(f"w:{name}") for name in local_names}
    for index, child in enumerate(parent):
        if child.tag in tags:
            parent.insert(index, element)
            return
    parent.append(element)


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        insert_before_any(
            tc_pr,
            shd,
            ("noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark"),
        )
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), fill)


def set_cell_borders(cell, **kwargs):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        insert_before_any(
            tc_pr,
            borders,
            ("shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark"),
        )
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        if edge not in kwargs:
            continue
        element = borders.find(qn(f"w:{edge}"))
        if element is None:
            element = OxmlElement(f"w:{edge}")
            borders.append(element)
        for key, value in kwargs[edge].items():
            element.set(qn(f"w:{key}"), str(value))


def set_table_borders(table, color="000000", size="8"):
    borders = table._tbl.tblPr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        insert_before_any(
            table._tbl.tblPr,
            borders,
            ("shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription"),
        )
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = borders.find(qn(f"w:{edge}"))
        if element is None:
            element = OxmlElement(f"w:{edge}")
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), size)
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), color)


def set_fixed_table_grid(table, grid_widths):
    """Lock table, grid and cell widths so WPS cannot reflow the outer edges."""
    grid_widths = [int(width) for width in grid_widths]
    total = sum(grid_widths)
    tbl_pr = table._tbl.tblPr

    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.insert(0, tbl_w)
    tbl_w.set(qn("w:type"), "dxa")
    tbl_w.set(qn("w:w"), str(total))

    layout = tbl_pr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        insert_before_any(tbl_pr, layout, ("tblCellMar", "tblLook", "tblCaption", "tblDescription"))
    layout.set(qn("w:type"), "fixed")

    indent = tbl_pr.find(qn("w:tblInd"))
    if indent is not None:
        tbl_pr.remove(indent)

    grid = table._tbl.tblGrid
    columns = grid.findall(qn("w:gridCol"))
    while len(columns) > len(grid_widths):
        grid.remove(columns[-1])
        columns = grid.findall(qn("w:gridCol"))
    while len(columns) < len(grid_widths):
        grid.append(OxmlElement("w:gridCol"))
        columns = grid.findall(qn("w:gridCol"))
    for column, width in zip(columns, grid_widths):
        column.set(qn("w:w"), str(width))

    for row in table._tbl.findall(qn("w:tr")):
        column_index = 0
        for cell in row.findall(qn("w:tc")):
            tc_pr = cell.get_or_add_tcPr()
            span_node = tc_pr.find(qn("w:gridSpan"))
            span = int(span_node.get(qn("w:val"), "1")) if span_node is not None else 1
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.insert(0, tc_w)
            tc_w.set(qn("w:type"), "dxa")
            tc_w.set(qn("w:w"), str(sum(grid_widths[column_index : column_index + span])))
            column_index += span


def set_cell_margins(cell, top=120, start=120, bottom=120, end=120):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        insert_before_any(
            tc_pr,
            tc_mar,
            ("textDirection", "tcFitText", "vAlign", "hideMark"),
        )
    for margin, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{margin}"))
        if node is None:
            node = OxmlElement(f"w:{margin}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_run_font(run, name="宋体", size=12, bold=False, color=None):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def add_para(container, text="", size=12, bold=False, align=None, before=0, after=6):
    p = container.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    if align is not None:
        p.alignment = align
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold)
    return p


def add_header_footer(section):
    section.header_distance = Cm(0.5)
    section.footer_distance = Cm(0.4)

    header = section.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hp.paragraph_format.space_after = Pt(0)
    hp.add_run().add_picture(str(ASSETS / "header_strip.jpg"), width=Cm(17.8))

    footer = section.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.paragraph_format.space_before = Pt(0)
    fp.add_run().add_picture(str(ASSETS / "footer_strip.jpg"), width=Cm(18.8))


def build_docx(data, output):
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.1)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(1.7)
    section.right_margin = Cm(1.7)
    add_header_footer(section)

    normal = doc.styles["Normal"]
    normal.font.name = "宋体"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    normal.font.size = Pt(12)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_before = Pt(18)
    title.paragraph_format.space_after = Pt(20)
    set_run_font(title.add_run(data.get("title", "会议纪要")), "黑体", 20, True)

    # A single 12-column master table is the WPS-stable layout primitive.
    # 12 columns can express the 4-column info area and 3-column sign-off area
    # without creating separate top-level tables that WPS may reflow differently.
    master = doc.add_table(rows=6, cols=12)
    master.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_table_borders(master)

    quarter = []
    for row_no in (0, 1):
        quarter.append(
            [
                master.cell(row_no, start).merge(master.cell(row_no, start + 2))
                for start in (0, 3, 6, 9)
            ]
        )
    attendee_cells = [
        master.cell(2, 0).merge(master.cell(2, 2)),
        master.cell(2, 3).merge(master.cell(2, 11)),
    ]
    body_cell = master.cell(3, 0).merge(master.cell(3, 11))
    sign_cells = []
    for row_no in (4, 5):
        sign_cells.append(
            [
                master.cell(row_no, start).merge(master.cell(row_no, start + 3))
                for start in (0, 4, 8)
            ]
        )

    # 9976 dxa matches the established company template width. Each quarter is
    # exactly 2494 dxa; the full grid is shared by every zone of the document.
    master_grid = [831, 831, 832] * 4
    set_fixed_table_grid(master, master_grid)

    for row in master.rows[:3]:
        seen = set()
        for cell in row.cells:
            if id(cell._tc) in seen:
                continue
            seen.add(id(cell._tc))
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_margins(cell, 140, 140, 140, 140)

    rows = [
        ("会议名称", data.get("meeting_name", "待补充"), "地点", data.get("place", "待补充")),
        ("主持人", data.get("host", "待补充"), "会议时间", data.get("meeting_time", "待补充")),
    ]
    for ridx, values in enumerate(rows):
        for cidx, value in enumerate(values):
            cell = quarter[ridx][cidx]
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if cidx in (0, 2) else WD_ALIGN_PARAGRAPH.LEFT
            set_run_font(p.add_run(value), "黑体" if cidx in (0, 2) else "宋体", 12, cidx in (0, 2))
            if cidx in (0, 2):
                set_cell_shading(cell, "F2F2F2")

    for cidx, value in enumerate(("参会人员", data.get("attendees", "待补充"))):
        cell = attendee_cells[cidx]
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if cidx == 0 else WD_ALIGN_PARAGRAPH.LEFT
        set_run_font(p.add_run(value), "黑体" if cidx == 0 else "宋体", 12, cidx == 0)
        if cidx == 0:
            set_cell_shading(cell, "F2F2F2")

    set_cell_margins(body_cell, top=520, start=180, bottom=360, end=180)
    set_cell_borders(
        body_cell,
        top={"val": "single", "sz": "8", "space": "0", "color": "000000"},
        bottom={"val": "single", "sz": "8", "space": "0", "color": "000000"},
        left={"val": "single", "sz": "8", "space": "0", "color": "000000"},
        right={"val": "single", "sz": "8", "space": "0", "color": "000000"},
    )

    counter = 1
    for section_data in data.get("sections", []):
        heading = section_data.get("heading", "").strip()
        if heading:
            add_para(body_cell, heading, size=14, bold=True, before=5, after=8)
        for item in section_data.get("items", []):
            item = str(item).strip()
            if not item:
                continue
            prefix = "" if item[:2].rstrip("、.").isdigit() else f"{counter}、"
            add_para(body_cell, f"{prefix}{item}", size=11.5, after=6)
            counter += 1

    for row in sign_cells:
        for cell in row:
            set_cell_borders(
                cell,
                top={"val": "nil"},
                bottom={"val": "nil"},
                left={"val": "nil"},
                right={"val": "nil"},
                insideH={"val": "nil"},
                insideV={"val": "nil"},
            )
            set_cell_margins(cell, 80, 120, 80, 120)
    for idx, text in enumerate(["编制：", "审核：", "批准："]):
        set_run_font(sign_cells[0][idx].paragraphs[0].add_run(text), "黑体", 13, True)
    set_run_font(sign_cells[1][0].paragraphs[0].add_run(f"抄送：{data.get('cc', '相关部门')}"), "黑体", 13, True)

    # Merges can rewrite tcW values; normalize once more after all content and
    # cell-level border overrides have been applied.
    set_fixed_table_grid(master, master_grid)

    doc.save(output)


def main():
    parser = argparse.ArgumentParser(description="Build NPA standard meeting minutes DOCX from JSON.")
    parser.add_argument("--input", required=True, help="Path to minutes JSON.")
    parser.add_argument("--output", required=True, help="Path to output DOCX.")
    args = parser.parse_args()

    with open(args.input, "r", encoding="utf-8") as f:
        data = json.load(f)
    build_docx(data, args.output)


if __name__ == "__main__":
    main()
