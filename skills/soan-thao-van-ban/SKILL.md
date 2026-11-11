---
name: soan-thao-van-ban
description: >-
  Quy trình và bộ biểu mẫu chuẩn để soạn thảo công văn, quyết định, nghị quyết, quy chế, quy định,
  thông báo, kế hoạch, báo cáo, tờ trình, biên bản, giấy mời, giấy giới thiệu, giấy nghỉ phép, phụ lục
  và văn bản quản lý nội bộ theo Nghị định 30/2020/NĐ-CP, Quyết định 1488/QĐ-ĐHYDCT và Quyết định
  1494/QĐ-ĐHYDCT của Trường Đại học Y Dược Cần Thơ. Kích hoạt khi người dùng yêu cầu soạn thảo, rà soát
  thể thức hoặc xuất file Word (.docx) cho bất kỳ văn bản hành chính hay văn bản quản lý nội bộ nào.
---

# Soạn Thảo Văn Bản Hành Chính & Quản Lý Nội Bộ Chuẩn

Skill này chuẩn hóa toàn bộ quy trình tiếp nhận yêu cầu, kiểm tra thông tin còn thiếu, soạn thảo đúng mẫu và xuất tệp Word (`.docx`) chuẩn thể thức theo:
- **Nghị định số 30/2020/NĐ-CP** (Công tác văn thư).
- **Quyết định số 1488/QĐ-ĐHYDCT** ngày 31/7/2020 (Quy định công tác văn thư – 12 biểu mẫu hành chính).
- **Quyết định số 1494/QĐ-ĐHYDCT** ngày 01/4/2026 (Quy chế xây dựng và ban hành VBQLNB – 10 biểu mẫu quản lý nội bộ).

---

## Tài Liệu Tra Cứu Trong Skill

| Tài liệu | Mục đích sử dụng |
| :--- | :--- |
| [the-thuc-va-ky-thuat.md](./references/the-thuc-va-ky-thuat.md) | Quy định khổ giấy A4, lề, phông chữ, cỡ chữ/kiểu chữ 14 thành phần thể thức, quy tắc ghi quyền hạn ký (`TM.`, `KT.`, `Q.`, `TL.`, `TUQ.`), ngày tháng, viện dẫn VB |
| [quy-tac-viet-hoa-va-viet-tat.md](./references/quy-tac-viet-hoa-va-viet-tat.md) | Quy tắc viết hoa tiếng Việt hành chính, bảng chữ viết tắt 25 loại VB + 3 loại bản sao, bảng ký hiệu đơn vị Trường ĐHYDCT |
| [bieu-mau-hanh-chinh.md](./references/bieu-mau-hanh-chinh.md) | Chi tiết 12 mẫu hành chính (**Mẫu 1 → Mẫu 11**) kèm checklist trường bắt buộc/tùy chọn |
| [bieu-mau-quan-ly-noi-bo.md](./references/bieu-mau-quan-ly-noi-bo.md) | Chi tiết 10 mẫu VBQLNB (**Mẫu VP.VBQLNB.01 → 10**) và quy trình 5 bước ban hành/sửa đổi/bãi bỏ Quy chế, Quy định |
| [generate_docx.py](./scripts/generate_docx.py) | Script Python tạo tệp `.docx` chuẩn A4, lề, Times New Roman, bảng header/chữ ký không viền, số trang ẩn trang 1 |

---

## Quy Trình Thực Hiện 5 Bước

### Bước 1: Định Tuyến Loại Văn Bản & Đọc Biểu Mẫu
1. Xác định loại văn bản người dùng yêu cầu và chọn đúng mẫu trong `references/`:
   - **Văn bản hành chính thông thường** → đọc [bieu-mau-hanh-chinh.md](./references/bieu-mau-hanh-chinh.md):
     - Nghị quyết (HĐT): **Mẫu 1** | Quyết định trực tiếp: **Mẫu 2** | QĐ ban hành/phê duyệt VB khác: **Mẫu 3a** + **Mẫu 3b**
     - Thông báo / Kế hoạch / Báo cáo / Tờ trình / Hướng dẫn / Chương trình / Đề án: **Mẫu 4**
     - Công văn: **Mẫu 5** | Giấy mời: **Mẫu 6** | Giấy giới thiệu: **Mẫu 7** | Biên bản: **Mẫu 8**
     - Giấy nghỉ phép: **Mẫu 9** | Phụ lục: **Mẫu 10** | Bản sao giấy (`SAO Y`, `SAO LỤC`, `TRÍCH SAO`): **Mẫu 11**
   - **Hồ sơ xây dựng/sửa đổi/thẩm định/bãi bỏ Quy chế, Quy định nội bộ** → đọc [bieu-mau-quan-ly-noi-bo.md](./references/bieu-mau-quan-ly-noi-bo.md):
     - **Mẫu VP.VBQLNB.01 → 10**.
2. Xác định **cấp ban hành** (để chọn cặp Cơ quan chủ quản / Đơn vị ban hành và ký hiệu số văn bản):
   - **Cấp Trường:** `BỘ Y TẾ` / `TRƯỜNG ĐẠI HỌC Y DƯỢC CẦN THƠ` (ký hiệu `.../QĐ-ĐHYDCT` hoặc `.../ĐHYDCT-[ĐƠN VỊ]`).
   - **Cấp Khoa / Phòng / Bộ môn** *(vd: Khoa Y, Bộ môn Mô phôi - Giải phẫu bệnh - Pháp Y)*: `TRƯỜNG ĐẠI HỌC Y DƯỢC CẦN THƠ` / `[TÊN KHOA/PHÒNG]` (hoặc `KHOA Y` / `BỘ MÔN ...`).

### Bước 2: Kiểm Tra Thông Tin & Chủ Động Hỏi Lại Khi Thiếu
1. Đối chiếu thông tin người dùng cung cấp với mục **"Kiểm tra thông tin (thiếu thì hỏi)"** của biểu mẫu đã chọn.
2. **Nếu thiếu bất kỳ thông tin bắt buộc nào**, **KHÔNG tự bịa** số liệu, tên người ký, thời gian, địa điểm hay căn cứ pháp lý chưa được cung cấp. Hãy **chủ động hỏi lại người dùng** ngắn gọn theo danh sách đánh số:
   - Chỉ hỏi những mục thực sự còn thiếu.
   - Đưa ra gợi ý mặc định hợp lý nếu có (ví dụ: *Địa danh mặc định: `Cần Thơ`*, *Đơn vị soạn thảo: `Bộ môn Mô phôi - Giải phẫu bệnh - Pháp Y` / `Khoa Y`*).
   - Nếu là **thông tin tùy chọn** thường chưa có khi mới soạn thảo (như *Số văn bản*, *Ngày ký chính thức*), mặc định để trống theo chuẩn `Số:       /...` và `Cần Thơ, ngày     tháng     năm 2026` mà không cần làm phiền người dùng trừ khi họ muốn điền sẵn.
3. Nếu người dùng yêu cầu *"cứ soạn khung trước, chỗ thiếu để trống"*, sử dụng dấu `...` hoặc `[cần bổ sung: ...]` đúng vị trí trong biểu mẫu.

### Bước 3: Soạn Thảo Chuẩn Thể Thức, Ngôn Ngữ & Viết Hoa
Kiểm tra lại toàn bộ bản thảo theo [the-thuc-va-ky-thuat.md](./references/the-thuc-va-ky-thuat.md) và [quy-tac-viet-hoa-va-viet-tat.md](./references/quy-tac-viet-hoa-va-viet-tat.md):
- **Ngày tháng:** Thêm `0` phía trước ngày `< 10` và tháng `1, 2` (vd: `ngày 05 tháng 01 năm 2026`).
- **Căn cứ ban hành:** In nghiêng; cuối mỗi căn cứ dùng `;`, cuối căn cứ cuối cùng dùng `.` (hoặc theo quy định riêng nếu có Luật/Nghị định sửa đổi, bổ sung).
- **Viện dẫn văn bản:** Lần đầu ghi đầy đủ loại, số, ký hiệu, ngày ban hành, cơ quan, tên gọi; viết hoa `Phần, Chương, Mục, Tiểu mục, Điều`, viết thường `khoản, điểm`.
- **Kết thúc văn bản:** Dấu `./.` ngay sau chữ cuối cùng của phần nội dung.
- **Khối chữ ký:** Đúng quyền hạn (`HIỆU TRƯỞNG`, `KT. HIỆU TRƯỞNG` / `PHÓ HIỆU TRƯỞNG`, `TM. HỘI ĐỒNG` / `CHỦ TỊCH`, `TL.`, `TUQ.`, `TRƯỞNG BỘ MÔN`...); **không** ghi học hàm, học vị trước họ tên người ký văn bản hành chính thông thường.
- **Nơi nhận:** Đúng chuẩn `- Như trên;` / `- Như Điều ...;` và dòng cuối `- Lưu: VT, [VIẾT TẮT ĐƠN VỊ] ([số bản]).`.

### Bước 4: Trình Bày & Xuất Tệp Word (`.docx`)
1. Trình bày bản xem trước rõ ràng trong câu trả lời để người dùng kiểm tra nội dung.
2. Khi người dùng muốn xuất file Word `.docx` (hoặc lưu thẳng vào hồ sơ):
   - Chuẩn bị file JSON cấu trúc văn bản và chạy [generate_docx.py](./scripts/generate_docx.py):
     ```powershell
     python "Y:\.agents\skills\soan-thao-van-ban\scripts\generate_docx.py" --input "<path_to_json>" --output "<path_to_docx>"
     ```
   - Kiểm tra cấu trúc các trường JSON hỗ trợ trong `generate_docx.py`:
     - Header: `parent_org`, `org`, `number_symbol`, `doc_subject_cv` (cho Công văn), `place_date`
     - Tiêu đề: `doc_type`, `subject`, `companion_note`, `authority_title`
     - Nội dung: `kinh_gui` (str hoặc list), `basis` (list), `operative_keyword`, `body` (list gồm chuỗi hoặc dict `article`, `heading`, `center_bold`, `center_title`, `label_value`, `table`)
     - Ký tên & Nơi nhận: `recipients` (list), `signer_authority`, `signer_title`, `signer_name`; hoặc `is_minutes: true` với `secretary_name`, `chair_title`, `chair_name` (cho Biên bản).
   - Ngoài ra, có thể mở đối chiếu trực tiếp 22 tệp `.docx` mẫu gốc trong `resources/01-quan-ly-noi-bo-qd1494/` và `resources/02-hanh-chinh-qd1488/` khi cần biểu mẫu có bố cục đặc thù phức tạp (như Mẫu 9 Giấy nghỉ phép có mục Xác nhận bên dưới, Mẫu 11 Bản sao).

### Bước 5: Quy Tắc Đặt Tên Tệp & Lưu Trữ
- Khi lưu file thành phẩm vào ổ `Y:\`, tuân thủ quy tắc đặt tên của `quy-tac-luu-tru-van-ban`:
  `[LOAI VB]-[SO VB]-[TRICH YEU NGAN KHONG DAU]-[NGAY YYYYMMDD].docx`
  *(Ví dụ: `CV-00-De-nghi-cap-kinh-phi-hoi-thao-20261110.docx`, `BB-05-Hop-Bo-mon-thang-11-20261110.docx`).*
- Nếu có tạo/lưu tệp mới vào `Y:\`, nhắc hoặc thực thi cập nhật mục lục `python "Y:\.agents\scripts\update_index.py"`.
