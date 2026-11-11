#!/usr/bin/env python3
"""
Generate standards-compliant Vietnamese administrative & internal-management .docx files
per Nghị định 30/2020/NĐ-CP, QĐ 1488/QĐ-ĐHYDCT, and QĐ 1494/QĐ-ĐHYDCT.
Zero external dependencies (pure Python stdlib).
"""
import argparse
import json
import os
import zipfile
from xml.sax.saxutils import escape

CONTENT_TYPES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""

RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>"""

SETTINGS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
</w:settings>"""

STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:eastAsia="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        <w:color w:val="000000"/>
        <w:sz w:val="28"/>
        <w:szCs w:val="28"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:before="0" w:after="120" w:line="276" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
</w:styles>"""

HEADER_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:jc w:val="center"/>
      <w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
    </w:pPr>
    <w:fldSimple w:instr="PAGE">
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Times New Roman" w:eastAsia="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
          <w:sz w:val="26"/>
          <w:szCs w:val="26"/>
        </w:rPr>
        <w:t>2</w:t>
      </w:r>
    </w:fldSimple>
  </w:p>
</w:hdr>"""


def run_xml(text, sz=14, bold=False, italic=False, color="000000"):
    sz_half = int(round(sz * 2))
    b = "<w:b/><w:bCs/>" if bold else ""
    i = "<w:i/><w:iCs/>" if italic else ""
    t = escape(text)
    return (
        f'<w:r><w:rPr>'
        f'<w:rFonts w:ascii="Times New Roman" w:eastAsia="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
        f'{b}{i}<w:color w:val="{color}"/><w:sz w:val="{sz_half}"/><w:szCs w:val="{sz_half}"/>'
        f'</w:rPr><w:t xml:space="preserve">{t}</w:t></w:r>'
    )


def p_xml(runs, align="both", indent_cm=0, before_pt=0, after_pt=6, line_mult=1.15, keep_next=False):
    jc = f'<w:jc w:val="{align}"/>' if align else ""
    ind = f'<w:ind w:firstLine="{int(round(indent_cm * 567))}"/>' if indent_cm else ""
    before = int(round(before_pt * 20))
    after = int(round(after_pt * 20))
    line = int(round(line_mult * 240))
    kn = "<w:keepNext/>" if keep_next else ""
    inner = "".join(runs) if isinstance(runs, list) else runs
    return (
        f'<w:p><w:pPr>{kn}<w:spacing w:before="{before}" w:after="{after}" w:line="{line}" w:lineRule="auto"/>'
        f'{ind}{jc}</w:pPr>{inner}</w:p>'
    )


def line_rule_p(length_chars=18, align="center", after_pt=4, sz=10):
    # Clean horizontal line using solid box-drawing / underscore or paragraph bottom border
    bar = "\u2500" * length_chars
    return p_xml(run_xml(bar, sz=sz, bold=False), align=align, before_pt=0, after_pt=after_pt, line_mult=1.0)


def borderless_table(rows_cells, col_widths_dxa):
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in col_widths_dxa)
    total_w = sum(col_widths_dxa)
    tbl_borders = (
        '<w:tblBorders>'
        '<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '</w:tblBorders>'
    )
    trs = []
    for row in rows_cells:
        tcs = []
        for idx, cell_ps in enumerate(row):
            w = col_widths_dxa[idx]
            content = "".join(cell_ps) if cell_ps else p_xml("", align="left", after_pt=0, line_mult=1.0)
            tcs.append(f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>{content}</w:tc>')
        trs.append(f'<w:tr>{"".join(tcs)}</w:tr>')
    return (
        f'<w:tbl><w:tblPr><w:tblW w:w="{total_w}" w:type="dxa"/><w:jc w:val="center"/>{tbl_borders}</w:tblPr>'
        f'<w:tblGrid>{grid}</w:tblGrid>{"".join(trs)}</w:tbl>'
    )


def bordered_table(headers, rows, col_widths_dxa):
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in col_widths_dxa)
    total_w = sum(col_widths_dxa)
    tbl_borders = (
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tblBorders>'
    )
    trs = []
    if headers:
        tcs = []
        for idx, h in enumerate(headers):
            w = col_widths_dxa[idx]
            p = p_xml(run_xml(str(h), sz=13, bold=True), align="center", before_pt=2, after_pt=2, line_mult=1.15)
            tcs.append(f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>{p}</w:tc>')
        trs.append(f'<w:tr>{"".join(tcs)}</w:tr>')
    for row in rows:
        tcs = []
        for idx, val in enumerate(row):
            w = col_widths_dxa[idx]
            lines = str(val).split("\n")
            ps = [p_xml(run_xml(ln, sz=13), align="left", before_pt=2, after_pt=2, line_mult=1.15) for ln in lines]
            tcs.append(f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>{"".join(ps)}</w:tc>')
        trs.append(f'<w:tr>{"".join(tcs)}</w:tr>')
    return (
        f'<w:tbl><w:tblPr><w:tblW w:w="{total_w}" w:type="dxa"/><w:jc w:val="center"/>{tbl_borders}</w:tblPr>'
        f'<w:tblGrid>{grid}</w:tblGrid>{"".join(trs)}</w:tbl>'
    )


def build_document_xml(data):
    elements = []
    landscape = bool(data.get("landscape", False))
    # Printable width: A4 portrait 11906 - left(1701=3cm) - right(1134=2cm) = 9071 dxa (~16.0 cm)
    # Landscape: 16838 - 1701 - 1134 = 14003 dxa
    printable_w = 14003 if landscape else 9071
    left_col_w = int(printable_w * 0.45)
    right_col_w = printable_w - left_col_w

    is_appendix_only = data.get("is_appendix_only", False)

    if not is_appendix_only:
        # 1. Header Table (Agency + National Motto)
        left_ps = []
        if data.get("parent_org"):
            left_ps.append(p_xml(run_xml(data["parent_org"].upper(), sz=12, bold=False), align="center", after_pt=0, line_mult=1.0))
        if data.get("org"):
            left_ps.append(p_xml(run_xml(data["org"].upper(), sz=12, bold=True), align="center", after_pt=0, line_mult=1.0))
            left_ps.append(line_rule_p(length_chars=14, align="center", after_pt=4, sz=10))
        if data.get("number_symbol"):
            left_ps.append(p_xml(run_xml(data["number_symbol"], sz=13, bold=False), align="center", before_pt=2, after_pt=2, line_mult=1.0))
        if data.get("doc_subject_cv"):
            left_ps.append(p_xml(run_xml(data["doc_subject_cv"], sz=12, bold=False), align="center", before_pt=4, after_pt=4, line_mult=1.05))

        right_ps = [
            p_xml(run_xml("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", sz=12, bold=True), align="center", after_pt=0, line_mult=1.0),
            p_xml(run_xml("Độc lập - Tự do - Hạnh phúc", sz=13, bold=True), align="center", after_pt=0, line_mult=1.0),
            line_rule_p(length_chars=24, align="center", after_pt=4, sz=10),
        ]
        if data.get("place_date"):
            right_ps.append(p_xml(run_xml(data["place_date"], sz=13, italic=True), align="center", before_pt=2, after_pt=4, line_mult=1.0))

        elements.append(borderless_table([[left_ps, right_ps]], [left_col_w, right_col_w]))

    # 2. Document Type & Subject (for typed documents & appendices)
    if data.get("doc_type"):
        elements.append(p_xml(run_xml(data["doc_type"].upper(), sz=14, bold=True), align="center", before_pt=10, after_pt=2, line_mult=1.1, keep_next=True))
    if data.get("subject"):
        for sline in str(data["subject"]).split("\n"):
            elements.append(p_xml(run_xml(sline, sz=14, bold=True), align="center", before_pt=0, after_pt=2, line_mult=1.1, keep_next=True))
        if not data.get("companion_note"):
            elements.append(line_rule_p(length_chars=18, align="center", after_pt=6, sz=10))
    if data.get("companion_note"):
        for cline in str(data["companion_note"]).split("\n"):
            elements.append(p_xml(run_xml(cline, sz=13, italic=True), align="center", before_pt=0, after_pt=2, line_mult=1.1))
        elements.append(line_rule_p(length_chars=18, align="center", after_pt=6, sz=10))

    # 3. Authority line (e.g. HIỆU TRƯỞNG TRƯỜNG ĐẠI HỌC Y DƯỢC CẦN THƠ)
    if data.get("authority_title"):
        elements.append(p_xml(run_xml(data["authority_title"].upper(), sz=14, bold=True), align="center", before_pt=6, after_pt=6, line_mult=1.15))

    # 4. Kính gửi (for Công văn, Tờ trình, Báo cáo, Giấy giới thiệu)
    if data.get("kinh_gui"):
        kg = data["kinh_gui"]
        if isinstance(kg, str):
            elements.append(p_xml(run_xml(f"Kính gửi: {kg}", sz=14, bold=False), align="center", before_pt=6, after_pt=6, line_mult=1.15))
        elif isinstance(kg, list):
            if len(kg) == 1:
                elements.append(p_xml(run_xml(f"Kính gửi: {kg[0]}", sz=14, bold=False), align="center", before_pt=6, after_pt=6, line_mult=1.15))
            else:
                elements.append(p_xml(run_xml("Kính gửi:", sz=14, bold=False), align="left", indent_cm=2.5, before_pt=6, after_pt=2, line_mult=1.15))
                for idx, item in enumerate(kg):
                    sep = "." if idx == len(kg) - 1 else ";"
                    text_item = item if item.endswith((";", ".")) else f"{item}{sep}"
                    elements.append(p_xml(run_xml(f"- {text_item}", sz=14, bold=False), align="left", indent_cm=4.0, before_pt=0, after_pt=2, line_mult=1.15))

    # 5. Legal basis lines (Căn cứ ...)
    for b_line in data.get("basis", []):
        elements.append(p_xml(run_xml(b_line, sz=14, italic=True), align="both", indent_cm=1.0, before_pt=0, after_pt=4, line_mult=1.15))

    # 6. Operative keyword (e.g. QUYẾT ĐỊNH: / QUYẾT NGHỊ:)
    if data.get("operative_keyword"):
        elements.append(p_xml(run_xml(data["operative_keyword"].upper(), sz=14, bold=True), align="center", before_pt=6, after_pt=6, line_mult=1.15))

    # 7. Body blocks
    for blk in data.get("body", []):
        if isinstance(blk, str):
            elements.append(p_xml(run_xml(blk, sz=14), align="both", indent_cm=1.0, before_pt=0, after_pt=6, line_mult=1.15))
        elif isinstance(blk, dict):
            btype = blk.get("type", "p")
            if btype == "center_bold":
                elements.append(p_xml(run_xml(blk["text"], sz=blk.get("sz", 14), bold=True), align="center", before_pt=6, after_pt=2, line_mult=1.15))
            elif btype == "center_title":
                elements.append(p_xml(run_xml(blk["text"], sz=blk.get("sz", 14), bold=True), align="center", before_pt=0, after_pt=6, line_mult=1.15))
            elif btype == "article":
                runs = [run_xml(blk["title"] + " ", sz=14, bold=True)]
                if blk.get("text"):
                    runs.append(run_xml(blk["text"], sz=14, bold=False))
                elements.append(p_xml(runs, align="both", indent_cm=1.0, before_pt=2, after_pt=6, line_mult=1.15))
            elif btype == "heading":
                elements.append(p_xml(run_xml(blk["text"], sz=blk.get("sz", 14), bold=True), align="both", indent_cm=blk.get("indent_cm", 0), before_pt=4, after_pt=4, line_mult=1.15))
            elif btype == "label_value":
                runs = [run_xml(blk["label"] + " ", sz=14, bold=blk.get("bold_label", False)), run_xml(blk["value"], sz=14, bold=False)]
                elements.append(p_xml(runs, align="both", indent_cm=blk.get("indent_cm", 1.0), before_pt=0, after_pt=4, line_mult=1.15))
            elif btype == "table":
                col_w = blk.get("col_widths")
                if not col_w:
                    n = len(blk.get("headers") or (blk["rows"][0] if blk.get("rows") else [1]))
                    col_w = [printable_w // n] * n
                elements.append(bordered_table(blk.get("headers", []), blk.get("rows", []), col_w))
            else:
                elements.append(p_xml(run_xml(blk.get("text", ""), sz=blk.get("sz", 14), bold=blk.get("bold", False), italic=blk.get("italic", False)), align=blk.get("align", "both"), indent_cm=blk.get("indent_cm", 1.0), before_pt=0, after_pt=6, line_mult=1.15))

    # 8. Signature / Recipients section
    if data.get("is_minutes"):
        # Biên bản: Thư ký (left) | Chủ tọa (right), then Nơi nhận below
        sec_ps = [
            p_xml(run_xml("THƯ KÝ CUỘC HỌP", sz=13, bold=True), align="center", before_pt=8, after_pt=2, line_mult=1.0),
            p_xml(run_xml("(Chữ ký)", sz=13, italic=True), align="center", after_pt=36, line_mult=1.0),
            p_xml(run_xml(data.get("secretary_name", "Họ và tên"), sz=14, bold=True), align="center", after_pt=4, line_mult=1.0),
        ]
        chr_ps = [
            p_xml(run_xml("CHỦ TỌA", sz=13, bold=True), align="center", before_pt=8, after_pt=2, line_mult=1.0),
        ]
        if data.get("chair_title"):
            chr_ps.append(p_xml(run_xml(data["chair_title"].upper(), sz=13, bold=True), align="center", after_pt=2, line_mult=1.0))
        chr_ps.append(p_xml(run_xml("(Chữ ký, dấu)", sz=13, italic=True), align="center", after_pt=36, line_mult=1.0))
        chr_ps.append(p_xml(run_xml(data.get("chair_name", "Họ và tên"), sz=14, bold=True), align="center", after_pt=4, line_mult=1.0))
        elements.append(borderless_table([[sec_ps, chr_ps]], [left_col_w, right_col_w]))
        if data.get("recipients"):
            elements.append(p_xml(run_xml("Nơi nhận:", sz=12, bold=True, italic=True), align="left", before_pt=6, after_pt=1, line_mult=1.0))
            for r in data["recipients"]:
                elements.append(p_xml(run_xml(r, sz=11), align="left", before_pt=0, after_pt=1, line_mult=1.0))
    elif not data.get("omit_signature_block"):
        rec_ps = []
        if data.get("recipients"):
            rec_ps.append(p_xml(run_xml("Nơi nhận:", sz=12, bold=True, italic=True), align="left", before_pt=8, after_pt=1, line_mult=1.0))
            for r in data["recipients"]:
                rec_ps.append(p_xml(run_xml(r, sz=11), align="left", before_pt=0, after_pt=1, line_mult=1.0))
        sig_ps = []
        if data.get("signer_authority"):
            sig_ps.append(p_xml(run_xml(data["signer_authority"].upper(), sz=13, bold=True), align="center", before_pt=8, after_pt=1, line_mult=1.0))
        if data.get("signer_title"):
            sig_ps.append(p_xml(run_xml(data["signer_title"].upper(), sz=13, bold=True), align="center", before_pt=(0 if data.get("signer_authority") else 8), after_pt=2, line_mult=1.0))
        if data.get("signer_title") or data.get("signer_authority"):
            sig_ps.append(p_xml(run_xml("(Chữ ký, dấu)" if not data.get("hide_sign_hint") else "", sz=12, italic=True), align="center", after_pt=36, line_mult=1.0))
        if data.get("signer_name"):
            sig_ps.append(p_xml(run_xml(data["signer_name"], sz=14, bold=True), align="center", after_pt=4, line_mult=1.0))
        if rec_ps or sig_ps:
            elements.append(borderless_table([[rec_ps, sig_ps]], [left_col_w, right_col_w]))

    # Page setup
    pg_w = 16838 if landscape else 11906
    pg_h = 11906 if landscape else 16838
    orient = ' w:orient="landscape"' if landscape else ""
    sect_pr = (
        f'<w:sectPr>'
        f'<w:headerReference w:type="default" r:id="rId2"/>'
        f'<w:titlePg/>'
        f'<w:pgSz w:w="{pg_w}" w:h="{pg_h}"{orient}/>'
        f'<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1701" w:header="567" w:footer="567" w:gutter="0"/>'
        f'</w:sectPr>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<w:body>{"".join(elements)}{sect_pr}</w:body></w:document>'
    )


def write_docx(data, out_path):
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    doc_xml = build_document_xml(data)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES_XML)
        z.writestr("_rels/.rels", RELS_XML)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS_XML)
        z.writestr("word/styles.xml", STYLES_XML)
        z.writestr("word/settings.xml", SETTINGS_XML)
        z.writestr("word/header1.xml", HEADER_XML)
        z.writestr("word/document.xml", doc_xml.encode("utf-8"))
    return out_path


def main():
    parser = argparse.ArgumentParser(description="Generate standard .docx from JSON spec")
    parser.add_argument("--input", required=True, help="Path to JSON input file")
    parser.add_argument("--output", required=True, help="Path to output .docx file")
    args = parser.parse_args()
    with open(args.input, "r", encoding="utf-8") as f:
        data = json.load(f)
    out = write_docx(data, args.output)
    print(f"Generated: {out}")


if __name__ == "__main__":
    main()
