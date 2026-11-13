"""Generate synthetic example reports to test dms_compiler.py.

Creates in ./examples:
  dms_template.docx   - DMS template: front page, revision/approval pages, TOC,
                        list of tables, headers/footers, manual chapters with
                        {{placeholders}}, a conditional chapter and the
                        {{SCADE_CHAPTERS}} marker.
  scade_report_A/B/C.docx - SCADE-like reports of different lengths (C has
                        chapter numbers typed by hand instead of auto-numbering).
  reports.xlsx        - batch sheet: one row per report to compile.
"""
import struct
import zlib
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_BREAK
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Pt, RGBColor, Inches
from openpyxl import Workbook

OUT = Path(__file__).parent / "examples"


# ---------------------------------------------------------------- helpers
def add_field(paragraph, instr, placeholder="Update field (F9)"):
    """Insert a complex field (begin / instr / separate / result / end)."""
    def fld(kind):
        r = paragraph.add_run()
        c = OxmlElement("w:fldChar")
        c.set(qn("w:fldCharType"), kind)
        r._r.append(c)
    fld("begin")
    r = paragraph.add_run()
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = f" {instr} "
    r._r.append(it)
    fld("separate")
    paragraph.add_run(placeholder)
    fld("end")


def page_break(doc):
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def add_heading_numbering(doc, abstract_id=90, num_id=90):
    """Link multilevel auto-numbering (1, 1.1, 1.1.1) to Heading 1-3."""
    numbering = doc.part.numbering_part.element
    lvls = ""
    for i in range(3):
        text = ".".join(f"%{k + 1}" for k in range(i + 1))
        lvls += (
            f'<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
            f'<w:pStyle w:val="Heading{i + 1}"/><w:lvlText w:val="{text}"/>'
            f'<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="{432 + 144 * i}" '
            f'w:hanging="{432 + 144 * i}"/></w:pPr></w:lvl>'
        )
    abstract = parse_xml(
        f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{abstract_id}">'
        f'<w:multiLevelType w:val="multilevel"/>{lvls}</w:abstractNum>'
    )
    num = parse_xml(
        f'<w:num {nsdecls("w")} w:numId="{num_id}">'
        f'<w:abstractNumId w:val="{abstract_id}"/></w:num>'
    )
    first_num = numbering.find(qn("w:num"))
    if first_num is not None:
        first_num.addprevious(abstract)
    else:
        numbering.append(abstract)
    numbering.append(num)
    for i in range(3):
        ppr = doc.styles[f"Heading {i + 1}"].element.get_or_add_pPr()
        numpr = parse_xml(
            f'<w:numPr {nsdecls("w")}><w:ilvl w:val="{i}"/>'
            f'<w:numId w:val="{num_id}"/></w:numPr>'
        )
        spacing = ppr.find(qn("w:spacing"))
        if spacing is not None:
            spacing.addprevious(numpr)
        else:
            ppr.insert(0, numpr)


def add_caption(doc, title):
    p = doc.add_paragraph(style="Caption")
    p.add_run("Table ")
    add_field(p, "SEQ Table \\* ARABIC", "1")
    p.add_run(f": {title}")
    return p


def add_bookmark(paragraph, name, bm_id):
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bm_id))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bm_id))
    paragraph._p.insert(1 if paragraph._p.pPr is not None else 0, start)
    paragraph._p.append(end)


def add_hyperlink(paragraph, text, url=None, anchor=None):
    link = OxmlElement("w:hyperlink")
    if url:
        rid = paragraph.part.relate_to(url, RT.HYPERLINK, is_external=True)
        link.set(qn("r:id"), rid)
    else:
        link.set(qn("w:anchor"), anchor)
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rpr.append(color)
    rpr.append(u)
    r.append(rpr)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    link.append(r)
    paragraph._p.append(link)


def make_png(path, rgb=(40, 110, 180), w=240, h=80):
    row = b"\x00" + bytes(rgb) * w
    raw = row * h

    def chunk(tag, data):
        c = struct.pack(">I", len(data)) + tag + data
        return c + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)

    png = (b"\x89PNG\r\n\x1a\n"
           + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
           + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))
    path.write_bytes(png)


def fill_table(table, header, rows):
    for i, h in enumerate(header):
        cell = table.rows[0].cells[i]
        cell.text = h
        cell.paragraphs[0].runs[0].bold = True
    for r, row in enumerate(rows, start=1):
        for c, val in enumerate(row):
            table.rows[r].cells[c].text = str(val)


# ---------------------------------------------------------------- template
def make_template(path):
    doc = Document()
    add_heading_numbering(doc)
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)
    for i in (1, 2, 3):
        doc.styles[f"Heading {i}"].font.color.rgb = RGBColor(0x1F, 0x3A, 0x5F)
    front = doc.styles.add_style("Front Heading", WD_STYLE_TYPE.PARAGRAPH)
    front.base_style = doc.styles["Normal"]
    front.font.size = Pt(16)
    front.font.bold = True

    sec = doc.sections[0]
    sec.header.paragraphs[0].text = "DMS Report  |  {{doc_number}}  |  {{system_name}}"
    fp = sec.footer.paragraphs[0]
    fp.add_run("Company Confidential  -  Page ")
    add_field(fp, "PAGE", "1")
    fp.add_run(" of ")
    add_field(fp, "NUMPAGES", "1")

    # Front page
    doc.add_paragraph("Design Management System Report", style="Title")
    doc.add_paragraph("{{system_name}}", style="Subtitle")
    t = doc.add_table(rows=4, cols=2, style="Table Grid")
    fill_table(t, ["Field", "Value"], [
        ["Document number", "{{doc_number}}"],
        ["Revision", "{{revision}}"],
        ["Author", "{{author}}"]])
    page_break(doc)

    # Pages before the table of contents
    doc.add_paragraph("Revision History", style="Front Heading")
    t = doc.add_table(rows=2, cols=3, style="Table Grid")
    fill_table(t, ["Rev", "Date", "Description"],
               [["{{revision}}", "{{issue_date}}", "First issue"]])
    page_break(doc)
    doc.add_paragraph("Approvals", style="Front Heading")
    t = doc.add_table(rows=3, cols=3, style="Table Grid")
    fill_table(t, ["Role", "Name", "Signature"],
               [["Prepared by", "{{author}}", ""], ["Approved by", "", ""]])
    page_break(doc)

    doc.add_paragraph("Table of Contents", style="Front Heading")
    add_field(doc.add_paragraph(), 'TOC \\o "1-3" \\h \\z \\u')
    page_break(doc)
    doc.add_paragraph("List of Tables", style="Front Heading")
    add_field(doc.add_paragraph(), 'TOC \\h \\z \\c "Table"')
    page_break(doc)

    # Manual chapter before the SCADE content
    doc.add_heading("Introduction", level=1)
    p = doc.add_paragraph("This report documents the software design of the ")
    # placeholder deliberately split across runs, as Word often does
    p.add_run("{{sys")
    p.add_run("tem_name}}").bold = True
    p.add_run(", a {{system_type}} with a maximum operating pressure of "
              "{{max_pressure}} bar.")
    doc.add_heading("Scope", level=2)
    doc.add_paragraph("The following chapters are generated from the SCADE model "
                      "of the {{system_name}} and reproduced without modification.")
    doc.add_heading("Reference Documents", level=2)
    t = doc.add_table(rows=3, cols=2, style="Table Grid")
    fill_table(t, ["Ref", "Title"], [
        ["[1]", "System Requirements Specification of the {{system_name}}"],
        ["[2]", "Software Development Plan"]])
    add_caption(doc, "Reference documents")

    # Conditional manual chapter
    doc.add_paragraph("{{#if has_redundancy}}")
    doc.add_heading("Redundancy Concept", level=1)
    doc.add_paragraph("The {{system_name}} uses {{channels}} redundant channels. "
                      "Each channel runs the same SCADE model.")
    doc.add_paragraph("{{/if}}")

    # Where the SCADE chapters go
    doc.add_paragraph("{{SCADE_CHAPTERS}}")

    # Manual chapter after the SCADE content
    doc.add_heading("Conclusions", level=1)
    doc.add_paragraph("The design of the {{system_name}} described above "
                      "satisfies the requirements listed in [1].")
    doc.save(path)


# ---------------------------------------------------------------- scade
def make_scade(path, system, operators, typed_numbers=False, image=None):
    doc = Document()
    if not typed_numbers:
        add_heading_numbering(doc, abstract_id=77, num_id=77)
    doc.styles["Normal"].font.name = "Arial"
    doc.styles["Normal"].font.size = Pt(9)
    for i in (1, 2, 3):
        st = doc.styles[f"Heading {i}"]
        st.font.name = "Arial"
        st.font.color.rgb = RGBColor(0xB0, 0x30, 0x30)
    code = doc.styles.add_style("SCADE Code", WD_STYLE_TYPE.PARAGRAPH)
    code.base_style = doc.styles["Normal"]
    code.font.name = "Consolas"
    code.font.size = Pt(8)
    doc.sections[0].header.paragraphs[0].text = f"SCADE Suite report - {system}"

    # Front matter that must be skipped
    doc.add_paragraph(f"SCADE Design Report - {system}", style="Title")
    doc.add_paragraph("Generated by SCADE Suite Reporter")
    page_break(doc)
    doc.add_paragraph("Table of Contents", style="TOC Heading")
    add_field(doc.add_paragraph(), 'TOC \\o "1-3" \\h \\z \\u')
    page_break(doc)

    bm = 100

    def heading(text, level, number):
        nonlocal bm
        label = f"{number} {text}" if typed_numbers else text
        h = doc.add_heading(label, level=level)
        bm += 1
        return h

    for ci, (op, n_sub) in enumerate(operators, start=1):
        h = heading(f"Operator {op}", 1, f"{ci}")
        add_bookmark(h, f"_Ref_op_{op}", bm)
        p = doc.add_paragraph(f"Operator {op} is part of package Control. See the ")
        add_hyperlink(p, "SCADE documentation", url="https://example.com/scade/" + op)
        p.add_run(" for the modelling rules.")
        if ci > 1:
            prev = operators[ci - 2][0]
            p = doc.add_paragraph("This operator is called by ")
            add_hyperlink(p, f"Operator {prev}", anchor=f"_Ref_op_{prev}")
            p.add_run(".")

        heading("Interface", 2, f"{ci}.1")
        t = doc.add_table(rows=4, cols=3, style="Light Grid Accent 1")
        fill_table(t, ["Name", "Type", "Direction"], [
            [f"{op}_in", "float32", "input"],
            [f"{op}_enable", "bool", "input"],
            [f"{op}_out", "float32", "output"]])
        add_caption(doc, f"Interface of operator {op}")

        heading("Description", 2, f"{ci}.2")
        doc.add_paragraph("The operator performs the following steps:")
        for step in ("Acquire and validate the inputs", "Compute the command",
                     "Saturate and publish the output"):
            doc.add_paragraph(step, style="List Bullet")
        doc.add_paragraph(f"{op}_out = if {op}_enable then limit({op}_in) else 0.0;",
                          style="SCADE Code")
        if image and ci == 1:
            doc.add_picture(str(image), width=Inches(2.5))

        heading("Local Variables", 2, f"{ci}.3")
        for si in range(1, n_sub + 1):
            heading(f"Variable L{si}", 3, f"{ci}.3.{si}")
            t = doc.add_table(rows=2, cols=2, style="Table Grid")
            fill_table(t, ["Property", "Value"], [["Type", "float32"]])
            add_caption(doc, f"Properties of {op}.L{si}")
    doc.save(path)


def main():
    OUT.mkdir(exist_ok=True)
    img = OUT / "diagram.png"
    make_png(img)
    make_template(OUT / "dms_template.docx")
    make_scade(OUT / "scade_report_A.docx", "Fuel Pump Controller",
               [("PumpCtrl", 1), ("ValveCmd", 2), ("Monitor", 1)], image=img)
    make_scade(OUT / "scade_report_B.docx", "Brake Control Unit",
               [("BrakeCmd", 2), ("AntiSkid", 3), ("WheelSpeed", 1),
                ("Autobrake", 2), ("BITE", 1), ("ModeLogic", 2)], image=img)
    make_scade(OUT / "scade_report_C.docx", "Cabin Pressure Controller",
               [("PressLoop", 1), ("OutflowValve", 1)], typed_numbers=True)
    img.unlink()

    wb = Workbook()
    ws = wb.active
    ws.title = "reports"
    ws.append(["scade_report", "output", "system_name", "system_type", "doc_number",
               "revision", "issue_date", "author", "max_pressure",
               "has_redundancy", "channels"])
    ws.append(["scade_report_A.docx", "DMS_FuelPump.docx", "Fuel Pump Controller",
               "single-channel controller", "DMS-0001", "A", "2026-11-13",
               "F. Branca", 12, "no", ""])
    ws.append(["scade_report_B.docx", "DMS_BrakeControl.docx", "Brake Control Unit",
               "dual-channel controller", "DMS-0002", "B", "2026-11-13",
               "F. Branca", 210, "yes", 2])
    ws.append(["scade_report_C.docx", "DMS_CabinPressure.docx",
               "Cabin Pressure Controller", "triplex controller", "DMS-0003", "A",
               "2026-11-13", "F. Branca", 0.8, "yes", 3])
    wb.save(OUT / "reports.xlsx")
    print("Examples written to", OUT)


if __name__ == "__main__":
    main()
