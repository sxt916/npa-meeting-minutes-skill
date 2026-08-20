import sys
from pathlib import Path
from zipfile import ZipFile

from lxml import etree


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def cell_span(cell):
    node = cell.find("w:tcPr/w:gridSpan", NS)
    return int(node.get(W + "val", "1")) if node is not None else 1


def table_metrics(table):
    width = table.find("w:tblPr/w:tblW", NS)
    layout = table.find("w:tblPr/w:tblLayout", NS)
    grid = table.findall("w:tblGrid/w:gridCol", NS)
    return {
        "width_type": width.get(W + "type") if width is not None else None,
        "width": int(width.get(W + "w", "0")) if width is not None else 0,
        "layout": layout.get(W + "type") if layout is not None else None,
        "grid": [int(col.get(W + "w", "0")) for col in grid],
    }


def check_cell_widths(table, grid):
    for row_no, row in enumerate(table.findall("w:tr", NS), 1):
        column = 0
        for cell_no, cell in enumerate(row.findall("w:tc", NS), 1):
            span = cell_span(cell)
            expected = sum(grid[column : column + span])
            tc_w = cell.find("w:tcPr/w:tcW", NS)
            actual = int(tc_w.get(W + "w", "0")) if tc_w is not None else 0
            width_type = tc_w.get(W + "type") if tc_w is not None else None
            if width_type != "dxa" or abs(actual - expected) > 5:
                raise SystemExit(
                    f"FAIL: master row {row_no} cell {cell_no} is not locked to grid "
                    f"(type={width_type}, width={actual}, expected={expected})"
                )
            column += span
        if column != len(grid):
            raise SystemExit(f"FAIL: master row {row_no} does not cover the full grid")


def check_nil_borders(cell, row_no, cell_no):
    borders = cell.find("w:tcPr/w:tcBorders", NS)
    if borders is None:
        raise SystemExit(f"FAIL: sign-off row {row_no} cell {cell_no} has no border override")
    for edge in ("top", "left", "bottom", "right"):
        node = borders.find(f"w:{edge}", NS)
        if node is None or node.get(W + "val") not in ("nil", "none"):
            raise SystemExit(
                f"FAIL: sign-off row {row_no} cell {cell_no} {edge} border is visible"
            )


def check(path: Path) -> None:
    with ZipFile(path) as archive:
        root = etree.fromstring(archive.read("word/document.xml"))

    body = root.find("w:body", NS)
    bordered_paragraphs = body.xpath("./w:p[w:pPr/w:pBdr]", namespaces=NS)
    tables = body.findall("w:tbl", NS)

    if bordered_paragraphs:
        raise SystemExit(
            f"FAIL: found {len(bordered_paragraphs)} top-level bordered paragraphs; "
            "use the master table body row instead"
        )
    if len(tables) != 1:
        raise SystemExit(
            f"FAIL: WPS-stable minutes require exactly one top-level master table; "
            f"found {len(tables)}"
        )

    master = tables[0]
    rows = master.findall("w:tr", NS)
    if len(rows) < 6:
        raise SystemExit("FAIL: master table must contain 3 info, 1 body, and 2 sign-off rows")

    metrics = table_metrics(master)
    if metrics["layout"] != "fixed" or metrics["width_type"] != "dxa":
        raise SystemExit("FAIL: master table must use fixed layout and dxa preferred width")
    if len(metrics["grid"]) != 12:
        raise SystemExit(
            f"FAIL: master table must use a 12-column shared grid; found {len(metrics['grid'])}"
        )
    if abs(metrics["width"] - sum(metrics["grid"])) > 5:
        raise SystemExit(
            f"FAIL: master preferred width {metrics['width']} does not match grid "
            f"{sum(metrics['grid'])}"
        )

    body_cells = rows[3].findall("w:tc", NS)
    if len(body_cells) != 1 or cell_span(body_cells[0]) != len(metrics["grid"]):
        raise SystemExit("FAIL: body must be one full-width cell in row 4 of the master table")
    if len(body_cells[0].findall("w:p", NS)) < 2:
        raise SystemExit("FAIL: body content is not contained in the continuous master cell")
    if body_cells[0].find("w:tcPr/w:tcBorders", NS) is None:
        raise SystemExit("FAIL: body cell has no explicit continuous outer border")

    check_cell_widths(master, metrics["grid"])
    for row_no, row in enumerate(rows[4:6], 5):
        for cell_no, cell in enumerate(row.findall("w:tc", NS), 1):
            check_nil_borders(cell, row_no, cell_no)

    print(
        "PASS: one fixed-width master table keeps info, body, and sign-off zones "
        "on the same WPS-stable boundaries"
    )


if __name__ == "__main__":
    check(Path(sys.argv[1]))
