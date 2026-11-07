"""
Script: export_docx.py
Mục đích: Chuyển đổi nội dung biên bản cuộc họp sang file Word (.docx) chuẩn thể thức
hành chính theo:
- Quyết định số 1488/QĐ-ĐHYDCT (Quy định công tác văn thư ĐH Y Dược Cần Thơ - Mẫu 8)
- Quyết định số 1494/QĐ-ĐHYDCT (Quy chế văn bản quản lý nội bộ ĐH Y Dược Cần Thơ - Phụ lục I)
- Nghị định số 30/2020/NĐ-CP của Chính phủ.

Hỗ trợ các cấp:
1. Trường Đại học Y Dược Cần Thơ (cấp Trường)
2. Khoa Y - Bộ môn Mô phôi - Giải phẫu bệnh - Pháp Y (cấp Khoa & Bộ môn)
3. Hỗ trợ 2 chữ ký (Thư ký - Chủ tọa/Trưởng Bộ môn) hoặc 3 chữ ký (Giao nhận, Kiểm tra sự cố thiết bị)
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
    """Thiết lập khoảng đệm (padding) trong ô bảng"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def remove_table_borders(table):
    """Ẩn viền bảng cho các khối căn chỉnh hành chính"""
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
    """Kẻ viền bảng chuẩn cho bảng danh sách hoặc Action Items"""
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
    
    # 1. Cấu hình lề trang theo QĐ 1494/QĐ-ĐHYDCT & Nghị định 30/2020/NĐ-CP
    # Lề trên: 20mm (0.79 in), Lề dưới: 20mm, Lề trái: 30mm (1.18 in), Lề phải: 20mm
    for section in doc.sections:
        section.page_width = Inches(8.27)    # A4: 210mm
        section.page_height = Inches(11.69)  # A4: 297mm
        section.top_margin = Inches(0.79)
        section.bottom_margin = Inches(0.79)
        section.left_margin = Inches(1.18)
        section.right_margin = Inches(0.79)

    # Cài đặt phông chữ mặc định Times New Roman
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(13)
    font.color.rgb = RGBColor(0, 0, 0)

    # 2. Khối Header hành chính (2 cột)
    header_table = doc.add_table(rows=1, cols=2)
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(header_table)

    cell_left = header_table.cell(0, 0)
    cell_right = header_table.cell(0, 1)
    
    cell_left.width = Inches(2.8)
    cell_right.width = Inches(3.5)

    # Cột trái: Cơ quan chủ quản / Trường / Khoa / Bộ môn
    p_cq = cell_left.paragraphs[0]
    p_cq.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cq.paragraph_format.line_spacing = 1.15
    p_cq.paragraph_format.space_after = Pt(0)
    
    cq_chu_quan = data.get('co_quan_chu_quan', '').strip().upper()
    ten_khoa = data.get('ten_khoa', '').strip().upper()
    ten_don_vi = data.get('ten_don_vi', 'TRƯỜNG ĐẠI HỌC Y DƯỢC CẦN THƠ').strip().upper()
    so_bb = data.get('so_bien_ban', '').strip()

    if cq_chu_quan:
        run_cq = p_cq.add_run(cq_chu_quan + '\n')
        run_cq.font.size = Pt(12)
        run_cq.font.bold = False

    if ten_khoa:
        run_k = p_cq.add_run(ten_khoa + '\n')
        run_k.font.size = Pt(12)
        run_k.font.bold = False

    run_dv = p_cq.add_run(ten_don_vi + '\n')
    run_dv.font.size = Pt(12)
    run_dv.font.bold = True

    # Đường kẻ dưới tên đơn vị ban hành (1/3 đến 1/2 độ dài)
    run_line_dv = p_cq.add_run('───────────\n')
    run_line_dv.font.size = Pt(8)
    run_line_dv.font.bold = True

    if so_bb:
        run_so = p_cq.add_run(f'Số: {so_bb}')
        run_so.font.size = Pt(12)
        run_so.font.italic = False

    # Cột phải: Quốc hiệu & Tiêu ngữ & Địa danh
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

    dia_diem_ngay = data.get('dia_diem_ngay', 'Cần Thơ, ngày ... tháng ... năm 20...').strip()
    if dia_diem_ngay:
        run_dd = p_qh.add_run(dia_diem_ngay)
        run_dd.font.size = Pt(13)
        run_dd.font.italic = True

    # Khoảng cách trước Tên biên bản
    p_space = doc.add_paragraph()
    p_space.paragraph_format.space_before = Pt(8)
    p_space.paragraph_format.space_after = Pt(2)

    # 3. Tên loại văn bản & Trích yếu
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(4)
    run_title = p_title.add_run('BIÊN BẢN')
    run_title.font.size = Pt(14)
    run_title.font.bold = True

    sub_title = data.get('sub_title', '').strip()
    if sub_title:
        p_st = doc.add_paragraph()
        p_st.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_st.paragraph_format.space_after = Pt(2)
        run_st = p_st.add_run(sub_title.upper())
        run_st.font.size = Pt(13)
        run_st.font.bold = True

    trich_yeu = data.get('trich_yeu', '').strip()
    if trich_yeu:
        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_sub.paragraph_format.space_after = Pt(12)
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
                p_cnt.paragraph_format.first_line_indent = Inches(0.39) # thụt đầu dòng 1cm
                
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

    # 6. Khối Chữ ký (Hỗ trợ 2 cột hoặc 3 cột)
    signatures = data.get('signatures', None)
    if not signatures:
        # Mặc định 2 người ký (Thư ký & Chủ trì/Trưởng bộ môn)
        signatures = [
            {
                'role': data.get('chuc_danh_thu_ky', 'THƯ KÝ CUỘC HỌP'),
                'sub': '(Ký và ghi rõ họ tên)',
                'name': data.get('ten_thu_ky', 'KTV. Nguyễn Phúc Duy')
            },
            {
                'role': data.get('chuc_danh_chu_tri', 'CHỦ TRÌ CUỘC HỌP / TRƯỞNG BỘ MÔN'),
                'sub': '(Ký và ghi rõ họ tên)',
                'name': data.get('ten_chu_tri', 'TS.BS. Nguyễn Hồng Phong')
            }
        ]

    cols_count = len(signatures)
    sign_table = doc.add_table(rows=1, cols=cols_count)
    sign_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(sign_table)

    col_width = Inches(6.3 / cols_count)
    for idx, sig in enumerate(signatures):
        cell = sign_table.cell(0, idx)
        cell.width = col_width
        p_sig = cell.paragraphs[0]
        p_sig.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_sig.paragraph_format.line_spacing = 1.15
        
        # Chức danh
        r_role = p_sig.add_run(sig.get('role', '').upper() + '\n')
        r_role.font.size = Pt(12 if cols_count == 3 else 13)
        r_role.font.bold = True
        
        # Hướng dẫn ký
        r_sub = p_sig.add_run(sig.get('sub', '(Ký và ghi rõ họ tên)') + '\n\n\n\n\n')
        r_sub.font.size = Pt(11)
        r_sub.font.italic = True
        
        # Họ và tên
        r_name = p_sig.add_run(sig.get('name', ''))
        r_name.font.size = Pt(12 if cols_count == 3 else 13)
        r_name.font.bold = True

    # 7. Nơi nhận (nếu có)
    noi_nhan = data.get('noi_nhan', [])
    if noi_nhan:
        p_nn_title = doc.add_paragraph()
        p_nn_title.paragraph_format.space_before = Pt(14)
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
    parser = argparse.ArgumentParser(description="Xuất biên bản cuộc họp ra file docx chuẩn hành chính ĐHYD Cần Thơ & NĐ 30")
    parser.add_argument("--json", dest="json_path", help="Đường dẫn file dữ liệu JSON của biên bản")
    parser.add_argument("--out", dest="output_path", default="output_bien_ban.docx", help="Đường dẫn file .docx đầu ra")
    args = parser.parse_args()

    if args.json_path:
        with open(args.json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        saved = create_administrative_minutes(data, args.output_path)
        print(f"[OK] Xuat bien ban thanh cong: {saved}")
    else:
        # Dữ liệu mặc định mẫu: Cuộc họp Bộ môn Mô phôi - GPB - Pháp Y (Trường ĐHYD Cần Thơ)
        sample_data = {
            'co_quan_chu_quan': 'TRƯỜNG ĐẠI HỌC Y DƯỢC CẦN THƠ',
            'ten_khoa': 'KHOA Y',
            'ten_don_vi': 'BỘ MÔN MÔ PHÔI - GIẢI PHẪU BỆNH - PHÁP Y',
            'so_bien_ban': '15/BB-BMMPGPBPY',
            'dia_diem_ngay': 'Cần Thơ, ngày 07 tháng 11 năm 2026',
            'sub_title': 'CUỘC HỌP BỘ MÔN',
            'trich_yeu': 'V/v Giao ban chuyên môn định kỳ và triển khai kế hoạch thực tập mô học học kỳ 1',
            'sections': [
                {
                    'title': 'I. THỜI GIAN VÀ ĐỊA ĐIỂM',
                    'items': [
                        '- **Thời gian bắt đầu:** 08 giờ 30 phút, ngày 07 tháng 11 năm 2026.',
                        '- **Địa điểm:** Văn phòng Bộ môn Mô phôi - Giải phẫu bệnh - Pháp Y (Khoa Y, Trường ĐHYD Cần Thơ).',
                        '- **Hình thức họp:** Trực tiếp.'
                    ]
                },
                {
                    'title': 'II. THÀNH PHẦN THAM DỰ',
                    'items': [
                        '- **Chủ trì:** TS.BS. Nguyễn Hồng Phong – Trưởng Bộ môn.',
                        '- **Thư ký:** KTV. Nguyễn Phúc Duy – Kỹ thuật viên Bộ môn.',
                        '- **Thành viên có mặt:** 08/08 cán bộ, giảng viên, kỹ thuật viên trong Bộ môn (100%).',
                        '- **Vắng mặt:** 0.'
                    ]
                },
                {
                    'title': 'III. NỘI DUNG CUỘC HỌP',
                    'items': [
                        '1. **Đánh giá hoạt động chuyên môn và đào tạo:** Tiến độ giảng dạy lý thuyết và thực hành giải phẫu bệnh cho các lớp Y đa khoa, Răng Hàm Mặt diễn ra đúng kế hoạch. Hệ thống kính hiển vi thực tập được bảo dưỡng tốt.',
                        '2. **Tình hình vận hành Labo và vật tư:** KTV. Duy báo cáo hệ thống máy đúc mô và máy cắt Microtome hoạt động ổn định. Đã kiểm nhập đợt hóa chất cồn, formalin, xylen và parafin đủ dùng cho quý 4/2026.',
                        '3. **Nghiên cứu khoa học và phát triển kỹ thuật:** Thảo luận triển khai số hóa tiêu bản (WSI) phục vụ ngân hàng dữ liệu mô học trực tuyến.'
                    ]
                },
                {
                    'title': 'IV. KẾT LUẬN CỦA CHỦ TRÌ (TRƯỞNG BỘ MÔN)',
                    'items': [
                        '- TS.BS. Nguyễn Hồng Phong biểu dương tinh thần làm việc của toàn thể viên chức.',
                        '- Yêu cầu chuẩn bị chu đáo bộ tiêu bản mẫu cho đợt thi kết thúc học phần thực hành sắp tới.',
                        '- **Phân công nhiệm vụ cụ thể (Action Items):**'
                    ],
                    'table': [
                        ['STT', 'Nội dung nhiệm vụ', 'Người thực hiện', 'Người/Đơn vị phối hợp', 'Thời hạn'],
                        ['1', 'Rà soát và chuẩn bị bộ lam kính thực hành', 'KTV. Nguyễn Phúc Duy', 'Giảng viên hướng dẫn', '12/11/2026'],
                        ['2', 'Kiểm tra nồng độ formalin và an toàn phòng Lab', 'KTV. Nguyễn Phúc Duy', 'Tổ Hóa chất', '10/11/2026'],
                        ['3', 'Tổng hợp lịch thi thực tập và danh sách sinh viên', 'Thư ký giáo vụ', 'Phòng Đào tạo Đại học', '15/11/2026']
                    ]
                }
            ],
            'ket_thuc': 'Cuộc họp kết thúc vào lúc 10 giờ 30 phút cùng ngày. Biên bản đã được thông qua toàn thể cuộc họp và thống nhất lưu hồ sơ./.',
            'signatures': [
                {
                    'role': 'THƯ KÝ CUỘC HỌP',
                    'sub': '(Ký và ghi rõ họ tên)',
                    'name': 'KTV. Nguyễn Phúc Duy'
                },
                {
                    'role': 'CHỦ TRÌ CUỘC HỌP / TRƯỞNG BỘ MÔN',
                    'sub': '(Ký và ghi rõ họ tên)',
                    'name': 'TS.BS. Nguyễn Hồng Phong'
                }
            ],
            'noi_nhan': [
                'Như thành phần tham dự;',
                'Ban Chủ nhiệm Khoa Y (để b/c);',
                'Lưu: BM, hồ sơ.'
            ]
        }
        target_path = os.path.join(os.path.dirname(__file__), "..", "examples", "sample_output.docx")
        saved = create_administrative_minutes(sample_data, target_path)
        print(f"[OK] Mau bien ban docx DH Y Duoc Can Tho da duoc tao: {saved}")

if __name__ == '__main__':
    main()
