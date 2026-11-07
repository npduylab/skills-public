"""
Script: export_docx.py
Mục đích: Chuyển đổi nội dung biên bản cuộc họp (từ JSON hoặc cấu trúc dữ liệu)
sang file Word (.docx) chuẩn thể thức hành chính Việt Nam theo Nghị định 30/2020/NĐ-CP.

Quy chuẩn định dạng:
- Khổ giấy: A4 (210 x 297 mm)
- Căn lề (Margins):
  + Lề trên: 20 mm (0.79 inch)
  + Lề dưới: 20 mm (0.79 inch)
  + Lề trái: 30 mm (1.18 inch)
  + Lề phải: 20 mm (0.79 inch)
- Phông chữ (Font): Times New Roman, Unicode tiếng Việt
- Quốc hiệu / Tiêu ngữ & Đơn vị: Bảng không viền 2 cột căn chỉnh chuẩn
- Chữ ký Thư ký / Chủ tọa: Bảng không viền 2 cột cuối văn bản
"""

import sys
import os
import json
import argparse
import re
from pathlib import Path

# Đảm bảo in UTF-8 trên Windows console
if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_margins(cell, top=100, bottom=100, left=100, right=100):
    """Thiết lập padding cho ô trong bảng"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def remove_table_borders(table):
    """Xóa toàn bộ viền bảng để làm bố cục trình bày hành chính"""
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="none"/>\n'
        f'  <w:left w:val="none"/>\n'
        f'  <w:bottom w:val="none"/>\n'
        f'  <w:right w:val="none"/>\n'
        f'  <w:insideH w:val="none"/>\n'
        f'  <w:insideV w:val="none"/>\n'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

def style_data_table(table):
    """Tạo kiểu cho bảng biểu dữ liệu (bảng phân công công việc)"""
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
        f'  <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
        f'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
        f'  <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
        f'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>\n'
        f'  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>\n'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

def create_administrative_minutes(data: dict, output_path: str):
    doc = Document()
    
    # 1. Cấu hình trang & lề chuẩn Nghị định 30/2020/NĐ-CP
    sections = doc.sections
    for section in sections:
        section.page_width = Inches(8.27)    # A4: 210mm
        section.page_height = Inches(11.69)  # A4: 297mm
        section.top_margin = Inches(0.79)    # 20mm
        section.bottom_margin = Inches(0.79) # 20mm
        section.left_margin = Inches(1.18)   # 30mm
        section.right_margin = Inches(0.79)  # 20mm

    # Normal style
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(13)
    font.color.rgb = RGBColor(0, 0, 0)

    # 2. Khối Tiêu đề Hành chính (Header Table)
    header_table = doc.add_table(rows=1, cols=2)
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(header_table)

    cell_left = header_table.cell(0, 0)
    cell_right = header_table.cell(0, 1)
    
    cell_left.width = Inches(2.7)
    cell_right.width = Inches(3.6)

    # Cột trái: Tên cơ quan ban hành
    p_cq = cell_left.paragraphs[0]
    p_cq.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cq.paragraph_format.line_spacing = 1.15
    p_cq.paragraph_format.space_after = Pt(0)
    
    cq_chu_quan = data.get('co_quan_chu_quan', '').strip().upper()
    ten_don_vi = data.get('ten_don_vi', '').strip().upper()
    so_bb = data.get('so_bien_ban', '').strip()

    if cq_chu_quan:
        run_cq = p_cq.add_run(cq_chu_quan + '\n')
        run_cq.font.size = Pt(12)
        run_cq.font.bold = False

    run_dv = p_cq.add_run(ten_don_vi + '\n')
    run_dv.font.size = Pt(12)
    run_dv.font.bold = True

    # Đường kẻ dưới tên cơ quan
    run_line_dv = p_cq.add_run('───────────\n')
    run_line_dv.font.size = Pt(8)
    run_line_dv.font.bold = True

    if so_bb:
        run_so = p_cq.add_run(f'Số: {so_bb}')
        run_so.font.size = Pt(12)
        run_so.font.italic = False

    # Cột phải: Quốc hiệu & Tiêu ngữ
    p_qh = cell_right.paragraphs[0]
    p_qh.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_qh.paragraph_format.line_spacing = 1.15
    p_qh.paragraph_format.space_after = Pt(0)

    run_qh = p_qh.add_run('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n')
    run_qh.font.size = Pt(12)
    run_qh.font.bold = True

    run_tn = p_qh.add_run('Độc lập - Tự do - Hạnh phúc\n')
    run_tn.font.size = Pt(13)
    run_tn.font.bold = True

    run_line_tn = p_qh.add_run('─────────────────────────\n')
    run_line_tn.font.size = Pt(8)
    run_line_tn.font.bold = True

    dia_diem_ngay = data.get('dia_diem_ngay', '').strip()
    if dia_diem_ngay:
        run_dd = p_qh.add_run(dia_diem_ngay)
        run_dd.font.size = Pt(13)
        run_dd.font.italic = True

    # Khoảng cách trước Tên biên bản
    p_space = doc.add_paragraph()
    p_space.paragraph_format.space_before = Pt(10)
    p_space.paragraph_format.space_after = Pt(2)

    # 3. Tên loại văn bản & Trích yếu
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(4)
    run_title = p_title.add_run('BIÊN BẢN')
    run_title.font.size = Pt(14)
    run_title.font.bold = True

    trich_yeu = data.get('trich_yeu', '').strip()
    if trich_yeu:
        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_sub.paragraph_format.space_after = Pt(14)
        run_sub = p_sub.add_run(trich_yeu)
        run_sub.font.size = Pt(13)
        run_sub.font.bold = True

    # 4. Nội dung chi tiết các mục
    sections_content = data.get('sections', [])
    for sec in sections_content:
        title = sec.get('title', '')
        content = sec.get('content', '')
        items = sec.get('items', [])
        table_data = sec.get('table', None)

        if title:
            p_sec = doc.add_paragraph()
            p_sec.paragraph_format.space_before = Pt(8)
            p_sec.paragraph_format.space_after = Pt(4)
            p_sec.paragraph_format.line_spacing = 1.2
            run_sec = p_sec.add_run(title)
            run_sec.font.size = Pt(13)
            run_sec.font.bold = True

        if content:
            for line in content.split('\n'):
                line = line.strip()
                if not line:
                    continue
                p_cnt = doc.add_paragraph()
                p_cnt.paragraph_format.line_spacing = 1.2
                p_cnt.paragraph_format.space_after = Pt(4)
                p_cnt.paragraph_format.first_line_indent = Inches(0.39) # 1cm
                
                parts = re.split(r'(\*\*.*?\*\*)', line)
                for part in parts:
                    if part.startswith('**') and part.endswith('**'):
                        r = p_cnt.add_run(part[2:-2])
                        r.bold = True
                    else:
                        p_cnt.add_run(part)

        if items:
            for itm in items:
                p_itm = doc.add_paragraph()
                p_itm.paragraph_format.line_spacing = 1.2
                p_itm.paragraph_format.space_after = Pt(3)
                p_itm.paragraph_format.left_indent = Inches(0.3)
                
                parts = re.split(r'(\*\*.*?\*\*)', itm)
                for part in parts:
                    if part.startswith('**') and part.endswith('**'):
                        r = p_itm.add_run(part[2:-2])
                        r.bold = True
                    else:
                        p_itm.add_run(part)

        if table_data and len(table_data) > 0:
            num_rows = len(table_data)
            num_cols = len(table_data[0])
            tbl = doc.add_table(rows=num_rows, cols=num_cols)
            style_data_table(tbl)
            
            for r_idx, row in enumerate(table_data):
                for c_idx, cell_value in enumerate(row):
                    cell = tbl.cell(r_idx, c_idx)
                    set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
                    p_c = cell.paragraphs[0]
                    p_c.paragraph_format.line_spacing = 1.15
                    p_c.paragraph_format.space_after = Pt(2)
                    
                    if r_idx == 0:
                        p_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        r = p_c.add_run(str(cell_value))
                        r.bold = True
                        r.font.size = Pt(11)
                    else:
                        if c_idx == 0 or c_idx == len(row) - 1:
                            p_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        else:
                            p_c.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        r = p_c.add_run(str(cell_value))
                        r.font.size = Pt(11)
            
            p_after_tbl = doc.add_paragraph()
            p_after_tbl.paragraph_format.space_after = Pt(6)

    # 5. Khối Thời gian kết thúc & Thông qua biên bản
    ket_thuc = data.get('ket_thuc', '')
    if ket_thuc:
        p_kt = doc.add_paragraph()
        p_kt.paragraph_format.space_before = Pt(8)
        p_kt.paragraph_format.space_after = Pt(12)
        p_kt.paragraph_format.line_spacing = 1.2
        p_kt.paragraph_format.first_line_indent = Inches(0.39)
        run_kt = p_kt.add_run(ket_thuc)
        run_kt.font.size = Pt(13)
        run_kt.font.italic = True

    # 6. Khối Chữ ký (Thư ký bên trái, Chủ tọa bên phải)
    sign_table = doc.add_table(rows=1, cols=2)
    sign_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(sign_table)

    s_left = sign_table.cell(0, 0)
    s_right = sign_table.cell(0, 1)
    s_left.width = Inches(3.1)
    s_right.width = Inches(3.2)

    # Thư ký
    p_sl = s_left.paragraphs[0]
    p_sl.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sl.paragraph_format.line_spacing = 1.15
    run_thu_ky = p_sl.add_run('THƯ KÝ CUỘC HỌP\n')
    run_thu_ky.font.size = Pt(13)
    run_thu_ky.font.bold = True
    run_tk_sub = p_sl.add_run('(Ký và ghi rõ họ tên)\n\n\n\n\n')
    run_tk_sub.font.size = Pt(11)
    run_tk_sub.font.italic = True

    ten_thu_ky = data.get('ten_thu_ky', '')
    if ten_thu_ky:
        run_tk_name = p_sl.add_run(ten_thu_ky)
        run_tk_name.font.size = Pt(13)
        run_tk_name.font.bold = True

    # Chủ trì / Chủ tọa
    p_sr = s_right.paragraphs[0]
    p_sr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sr.paragraph_format.line_spacing = 1.15
    chuc_danh_chu_tri = data.get('chuc_danh_chu_tri', 'CHỦ TỌA / CHỦ TRÌ').upper()
    run_ct = p_sr.add_run(chuc_danh_chu_tri + '\n')
    run_ct.font.size = Pt(13)
    run_ct.font.bold = True
    run_ct_sub = p_sr.add_run('(Ký và ghi rõ họ tên)\n\n\n\n\n')
    run_ct_sub.font.size = Pt(11)
    run_ct_sub.font.italic = True

    ten_chu_tri = data.get('ten_chu_tri', '')
    if ten_chu_tri:
        run_ct_name = p_sr.add_run(ten_chu_tri)
        run_ct_name.font.size = Pt(13)
        run_ct_name.font.bold = True

    # 7. Nơi nhận (nếu có)
    noi_nhan = data.get('noi_nhan', [])
    if noi_nhan:
        p_nn_title = doc.add_paragraph()
        p_nn_title.paragraph_format.space_before = Pt(18)
        p_nn_title.paragraph_format.space_after = Pt(2)
        run_nn = p_nn_title.add_run('Nơi nhận:\n')
        run_nn.font.size = Pt(11)
        run_nn.font.bold = True
        run_nn.font.italic = True

        for nn in noi_nhan:
            p_nn_item = doc.add_paragraph()
            p_nn_item.paragraph_format.space_after = Pt(1)
            p_nn_item.paragraph_format.line_spacing = 1.0
            r_item = p_nn_item.add_run(f'- {nn}')
            r_item.font.size = Pt(10)

    # Lưu file
    output_p = Path(output_path)
    output_p.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_p))
    return str(output_p.resolve())

def main():
    parser = argparse.ArgumentParser(description="Xuất biên bản cuộc họp ra file docx chuẩn hành chính")
    parser.add_argument("--json", dest="json_path", help="Đường dẫn file dữ liệu JSON của biên bản")
    parser.add_argument("--out", dest="output_path", default="output_bien_ban.docx", help="Đường dẫn file .docx đầu ra")
    args = parser.parse_args()

    if args.json_path:
        with open(args.json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        saved = create_administrative_minutes(data, args.output_path)
        print(f"[OK] Xuat bien ban thanh cong: {saved}")
    else:
        # Dữ liệu mặc định mẫu để kiểm tra
        sample_data = {
            'co_quan_chu_quan': 'BỆNH VIỆN ĐA KHOA',
            'ten_don_vi': 'KHOA GIẢI PHẪU BỆNH',
            'so_bien_ban': '01/BB-KGPB',
            'dia_diem_ngay': 'Hà Nội, ngày 07 tháng 11 năm 2026',
            'trich_yeu': 'Cuộc họp giao ban chuyên môn định kỳ và triển khai nhiệm vụ công tác',
            'sections': [
                {
                    'title': 'I. THỜI GIAN VÀ ĐỊA ĐIỂM',
                    'items': [
                        '- **Thời gian:** 08 giờ 30 phút ngày 07/11/2026.',
                        '- **Địa điểm:** Văn phòng Khoa Giải phẫu bệnh.'
                    ]
                },
                {
                    'title': 'II. THÀNH PHẦN THAM DỰ',
                    'items': [
                        '- **Chủ tọa:** BS.CKII. Trần Văn A - Trưởng khoa.',
                        '- **Thư ký:** KTV. Nguyễn Phúc Duy - Kỹ thuật viên GPB.',
                        '- **Số lượng tham dự:** 12/12 cán bộ, nhân viên có mặt đầy đủ.'
                    ]
                },
                {
                    'title': 'III. NỘI DUNG VÀ DIỄN BIẾN CUỘC HỌP',
                    'content': 'Đồng chí Chủ tọa thông qua chương trình làm việc. Các bộ phận báo cáo tình hình hoạt động chuyên môn, tiến độ thực hiện kỹ thuật mô học, hóa mô miễn dịch và giải pháp khắc phục khó khăn.'
                },
                {
                    'title': 'IV. KẾT LUẬN CỦA CHỦ TỌA',
                    'items': [
                        '- Toàn thể cán bộ tuân thủ nghiêm ngặt quy trình kiểm soát chất lượng (QA/QC).',
                        '- Phân công các đầu việc cụ thể theo bảng bên dưới.'
                    ],
                    'table': [
                        ['STT', 'Nội dung công việc', 'Người thực hiện', 'Người phối hợp', 'Thời hạn'],
                        ['1', 'Rà soát quy trình nhuộm HE', 'KTV. Nguyễn Phúc Duy', 'Tổ Cắt nhuộm', '12/11/2026'],
                        ['2', 'Kiểm kê kho hóa chất, sinh phẩm', 'KTV. Lê Thị B', 'Kho Dược', '15/11/2026']
                    ]
                }
            ],
            'ket_thuc': 'Cuộc họp kết thúc hồi 10 giờ 30 phút cùng ngày. Biên bản đã được thông qua và nhất trí hoàn toàn.',
            'chuc_danh_chu_tri': 'TRƯỞNG KHOA',
            'ten_chu_tri': 'BS.CKII. Trần Văn A',
            'ten_thu_ky': 'KTV. Nguyễn Phúc Duy',
            'noi_nhan': [
                'Ban Giám đốc (để báo cáo);',
                'Các bộ phận trong khoa;',
                'Lưu: VT, Hồ sơ họp.'
            ]
        }
        target_path = os.path.join(os.path.dirname(__file__), "..", "examples", "sample_output.docx")
        saved = create_administrative_minutes(sample_data, target_path)
        print(f"[OK] Mau bien ban docx da duoc tao: {saved}")

if __name__ == '__main__':
    main()
