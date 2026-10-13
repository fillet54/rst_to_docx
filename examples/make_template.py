"""Build examples/template.docx, a sample SVD template.

Layout:
  page 1  instructions page: uses every style the renderer needs (so Word keeps
          them in styles.xml) and ends with {{end-instructions}}
  page 2  cover page with {{placeholders}}
  page 3  table of contents field
  then    {{content}}
Headers/footers carry placeholders too, and Heading 1-4 are numbered with a
multilevel list so the rendered sections come out as 1, 1.1, 1.1.1 ...

In practice you'd design the template in Word; this script just produces a
reproducible example.
"""

import sys
from pathlib import Path

import docx
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Pt, RGBColor

HEADING_LEVELS = 4


def add_missing_builtin_styles(doc):
    """Add built-in styles the python-docx default template lacks."""
    plain = doc.styles.add_style("Plain Text", WD_STYLE_TYPE.PARAGRAPH)
    plain.base_style = doc.styles["Normal"]
    plain.font.name = "Consolas"
    plain.font.size = Pt(9)
    plain.paragraph_format.space_after = Pt(6)

    link = doc.styles.add_style("Hyperlink", WD_STYLE_TYPE.CHARACTER)
    link.font.color.rgb = RGBColor(0x05, 0x63, 0xC1)
    link.font.underline = True

    code = doc.styles.add_style("HTML Code", WD_STYLE_TYPE.CHARACTER)
    code.font.name = "Consolas"
    code.font.size = Pt(9.5)


def number_headings(doc):
    """Attach a 1 / 1.1 / 1.1.1 multilevel list to Heading 1..N."""
    numbering = doc.part.numbering_part.element
    abstract_id = 900
    levels = []
    for i in range(HEADING_LEVELS):
        text = ".".join(f"%{n + 1}" for n in range(i + 1))
        indent = 432 + 144 * i
        levels.append(
            f'<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
            f'<w:pStyle w:val="Heading{i + 1}"/><w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/>'
            f'<w:pPr><w:ind w:left="{indent}" w:hanging="{indent}"/></w:pPr></w:lvl>'
        )
    abstract = parse_xml(
        f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{abstract_id}">'
        f'<w:multiLevelType w:val="multilevel"/>{"".join(levels)}</w:abstractNum>'
    )
    numbering.find(qn("w:num")).addprevious(abstract)  # abstractNums precede nums
    num_id = numbering.add_num(abstract_id).numId
    for i in range(HEADING_LEVELS):
        numPr = doc.styles[f"Heading {i + 1}"].element.get_or_add_pPr().get_or_add_numPr()
        numPr.get_or_add_ilvl().val = i
        numPr.get_or_add_numId().val = num_id
    # TOC Heading is based on Heading 1; numId 0 switches the numbering off.
    doc.styles["TOC Heading"].element.get_or_add_pPr().get_or_add_numPr().get_or_add_numId().val = 0


def add_field(paragraph, instr, placeholder=""):
    """Append a complex field (e.g. TOC, PAGE) showing ``placeholder`` until updated."""
    def field_run(child):
        paragraph._p.append(parse_xml(f"<w:r {nsdecls('w')}>{child}</w:r>"))

    field_run('<w:fldChar w:fldCharType="begin" w:dirty="true"/>')
    field_run(f'<w:instrText xml:space="preserve"> {instr} </w:instrText>')
    field_run('<w:fldChar w:fldCharType="separate"/>')
    paragraph.add_run(placeholder)
    field_run('<w:fldChar w:fldCharType="end"/>')


def instructions_page(doc):
    doc.add_paragraph("Template instructions (removed when rendering)", style="Title")
    doc.add_paragraph(
        "Everything up to and including the {{end-instructions}} paragraph below is "
        "deleted by rst_to_docx. Use this page to keep one example of every style the "
        "renderer uses: Word only saves a built-in style into the file once it has "
        "been used, and the renderer can only apply styles that are saved."
    )
    doc.add_paragraph(
        "Markers: put {{content}} alone in a paragraph where the RST body goes. "
        "Anywhere else (cover, headers, footers, tables) {{dotted.name}} is "
        "replaced from the variables dict."
    )
    for level in range(1, HEADING_LEVELS + 1):
        doc.add_paragraph(f"Heading {level} sample", style=f"Heading {level}")
    doc.add_paragraph("Normal paragraph sample.", style="Normal")
    for kind in ("List Bullet", "List Number", "List Continue"):
        for n in ("", " 2", " 3"):
            doc.add_paragraph(f"{kind}{n} sample", style=kind + n)
    doc.add_paragraph("Plain Text (literal block) sample", style="Plain Text")
    doc.add_paragraph("Caption (table title) sample", style="Caption")
    doc.add_paragraph("Quote sample", style="Quote")
    p = doc.add_paragraph("Character styles: ")
    p.add_run("Hyperlink", style="Hyperlink")
    p.add_run(" and ")
    p.add_run("HTML Code", style="HTML Code")
    table = doc.add_table(rows=2, cols=2, style="Table Grid")
    table.cell(0, 0).text = "Table Grid"
    table.cell(0, 1).text = "sample"
    doc.add_paragraph("{{end-instructions}}")
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def cover_page(doc):
    for _ in range(6):
        doc.add_paragraph()
    p = doc.add_paragraph("{{project}}", style="Title")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc.add_paragraph("Software Version Description", style="Subtitle")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    table = doc.add_table(rows=4, cols=2, style="Table Grid")
    rows = [
        ("Document number", "{{doc.number}}"),
        ("Software version", "{{release.version}}"),
        ("Release date", "{{release.date}}"),
        ("Prepared by", "{{doc.author}}"),
    ]
    for (label, value), row in zip(rows, table.rows):
        row.cells[0].text = label
        row.cells[1].text = value
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def toc_page(doc):
    doc.add_paragraph("Table of Contents", style="TOC Heading")
    add_field(doc.add_paragraph(), r'TOC \o "1-3" \h \z \u',
              "Right-click and choose Update Field to build the table of contents.")
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def header_footer(doc):
    section = doc.sections[0]
    section.header.paragraphs[0].text = "{{project}} SVD — {{doc.number}}"
    footer = section.footer.paragraphs[0]
    footer.text = "Version {{release.version}}\tPage "
    add_field(footer, "PAGE", "1")


def build(path):
    doc = docx.Document()
    add_missing_builtin_styles(doc)
    number_headings(doc)
    instructions_page(doc)
    cover_page(doc)
    toc_page(doc)
    doc.add_paragraph("{{content}}")
    header_footer(doc)
    doc.save(path)
    return path


if __name__ == "__main__":
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).with_name("template.docx")
    print(build(out))
