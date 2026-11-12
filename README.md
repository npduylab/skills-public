# NPDUYLAB Public AI Agent Skills Collection 🚀

Kho lưu trữ **Công khai (Public)** các bộ kỹ năng (Skills) dùng chung dành cho **AI Coding Assistants & Autonomous Agents** (Google Antigravity, Claude Code, Cursor, Codex...), được phát triển và tối ưu bởi **Nguyễn Phúc Duy (NPDUYLAB)**.

> 📌 *Các kỹ năng đặc thù nội bộ đơn vị và cấu hình cá nhân được quản lý tách biệt tại kho riêng tư (`skills-private`).*

---

## 📂 Danh mục Kỹ năng (Skills Catalog)

| Kỹ năng | Phiên bản | Mô tả tóm tắt | Quy chuẩn tích hợp | Nền tảng hỗ trợ |
| :--- | :---: | :--- | :--- | :--- |
| [`tao-bien-ban-cuoc-hop`](./skills/tao-bien-ban-cuoc-hop/SKILL.md) | `1.1.0` | Tự động chuyển đổi ghi chép thô thành **Biên bản cuộc họp chính thức** chuẩn thể thức hành chính. Hỗ trợ xuất đồng thời Markdown (.md) và Word (.docx). | **Đại học Y Dược Cần Thơ** (QĐ 1488/QĐ-ĐHYDCT & QĐ 1494/QĐ-ĐHYDCT), **Nghị định 30/2020/NĐ-CP** | Antigravity, Claude Code, Cursor, Codex |
| [`soan-thao-van-ban`](./skills/soan-thao-van-ban/SKILL.md) | `1.0.0` | Quy trình và bộ **22 biểu mẫu chuẩn** để soạn thảo công văn, quyết định, nghị quyết, quy chế, thông báo, kế hoạch, báo cáo, tờ trình, biên bản, giấy mời, giấy nghỉ phép... và xuất Word (`.docx`). | **Nghị định 30/2020/NĐ-CP**, **QĐ 1488/QĐ-ĐHYDCT** (12 mẫu hành chính) & **QĐ 1494/QĐ-ĐHYDCT** (10 mẫu VBQLNB) | Antigravity, Claude Code, Cursor, Codex |
| [`doc-van-ban-markitdown`](./skills/doc-van-ban-markitdown/SKILL.md) | `1.0.0` | Đọc và chuyển đổi tài liệu văn phòng (`.docx`, `.xlsx`, `.pptx`, `.pdf`, `.html`, `.zip`...) sang Markdown bằng **Microsoft MarkItDown** chuẩn UTF-8 trên Windows. | **Microsoft MarkItDown** (`markitdown[all]`), UTF-8 Windows PowerShell | Antigravity, Claude Code, Cursor, Codex |
| [`luu-kho-git`](./skills/luu-kho-git/SKILL.md) | `2.0.0` | Quy trình chuẩn hóa phân loại **Public/Private**, đồng bộ đa điểm qua Hub trung tâm (`d:\Antigravity_system`), rà soát rò rỉ bí mật và commit/push lên GitHub bằng `git` & `gh`. | **GitHub CLI (`gh`)**, quy trình 2 kho `skills-public` & `skills-private` | Antigravity, Claude Code, Cursor, Codex |

---

## 📌 Chi tiết Kỹ năng: `tao-bien-ban-cuoc-hop` (Phiên bản 1.1.0)

### 1. Giới thiệu & Cập nhật mới
Kỹ năng hướng dẫn AI Agent tiếp nhận các bản ghi chép nhanh, bullet points, transcript cuộc họp và biên soạn thành biên bản cuộc họp chính thức.

✨ **Điểm mới cập nhật phiên bản 1.1.0:**
- Tích hợp toàn diện biểu mẫu chuẩn của **Trường Đại học Y Dược Cần Thơ**:
  - **Quyết định số 1488/QĐ-ĐHYDCT:** Quy định về công tác văn thư (Áp dụng **Mẫu 8 - Biên bản** tại Phụ lục V).
  - **Quyết định số 1494/QĐ-ĐHYDCT:** Quy chế xây dựng và ban hành văn bản quản lý nội bộ (Phụ lục I: Thể thức, kỹ thuật trình bày văn bản).
  - Chuẩn hóa bộ mẫu thực tế của Khoa Y và **Bộ môn Mô phôi - Giải phẫu bệnh - Pháp Y**:
    + Mẫu BM05: Biên bản họp Bộ môn (giao ban định kỳ, phân công nhiệm vụ, kế hoạch đào tạo).
    + Mẫu BM04: Biên bản ghi nhận hiện trạng sự cố, hư hỏng trang thiết bị Labo.
    + Mẫu BM06: Biên bản giao nhận, kiểm nhập hóa chất & vật tư tiêu hao.
- Tự động điền đúng cấp cơ quan, địa danh (*Cần Thơ, ngày...*), số ký hiệu (`/BB-BMMPGPBPY`, `/BB-KY`, `/BB-ĐHYDCT`), thông tin chủ trì (*TS.BS. Nguyễn Hồng Phong*) và thư ký (*KTV. Nguyễn Phúc Duy*).
- Nâng cấp script Python `export_docx.py` hỗ trợ bố cục 2 chữ ký hoặc 3 chữ ký, căn lề chuẩn A4 (Trái 30mm, Phải 20mm, Trên 20mm, Dưới 20mm).

### 2. Cấu trúc thư mục kỹ năng
```text
skills/tao-bien-ban-cuoc-hop/
├── SKILL.md                                 # Hướng dẫn quy trình 5 bước thực hiện cho Agent
├── scripts/
│   └── export_docx.py                       # Script Python xuất file Word (.docx) chuẩn ĐHYD Cần Thơ & NĐ 30
├── references/
│   ├── qd-1488-1494-dhy-duoc-can-tho.md     # Bảng tra cứu quy chuẩn thể thức QĐ 1488 & QĐ 1494 ĐHYD Cần Thơ
│   ├── nghi-dinh-30-2020-nd-cp.md           # Bảng tra cứu quy chuẩn thể thức chung Nghị định 30/2020/NĐ-CP
│   └── templates.md                         # Bộ khung mẫu biên bản BM05, BM04, BM06
└── examples/
    ├── ghi_chu_mau.md                       # Dữ liệu ghi chú mẫu đầu vào (Giao ban Bộ môn Mô phôi - GPB - Pháp Y)
    ├── bien_ban_chuan_mau.md                # Biên bản hoàn chỉnh mẫu đầu ra (.md)
    └── sample_output.docx                   # File Word thành phẩm mẫu (.docx)
```

---

## 📌 Chi tiết Kỹ năng: `soan-thao-van-ban` (Phiên bản 1.0.0)

### 1. Giới thiệu
Bộ kỹ năng toàn diện cho phép AI Agent tiếp nhận yêu cầu soạn thảo, rà soát thể thức, chủ động hỏi lại thông tin còn thiếu và xuất file Word (`.docx`) chuẩn hành chính & quản lý nội bộ theo:
- **Nghị định 30/2020/NĐ-CP** về công tác văn thư.
- **Quyết định số 1488/QĐ-ĐHYDCT** (12 biểu mẫu hành chính: Mẫu 1 → Mẫu 11).
- **Quyết định số 1494/QĐ-ĐHYDCT** (10 biểu mẫu văn bản quản lý nội bộ: Mẫu VP.VBQLNB.01 → 10).
- Kèm script Python [`generate_docx.py`](./skills/soan-thao-van-ban/scripts/generate_docx.py) không cần thư viện ngoài (pure Python stdlib) và 22 tệp `.docx` mẫu gốc.

### 2. Cấu trúc thư mục kỹ năng
```text
skills/soan-thao-van-ban/
├── SKILL.md                                 # Hướng dẫn quy trình 5 bước thực hiện cho Agent
├── scripts/
│   └── generate_docx.py                     # Script Python tạo file Word (.docx) chuẩn A4, NĐ 30
├── references/
│   ├── the-thuc-va-ky-thuat.md              # Quy chuẩn khổ giấy, lề, font, 14 thành phần thể thức
│   ├── quy-tac-viet-hoa-va-viet-tat.md      # Quy tắc viết hoa, viết tắt 25 loại VB & ký hiệu đơn vị
│   ├── bieu-mau-hanh-chinh.md               # Chi tiết 12 mẫu hành chính (QĐ 1488/QĐ-ĐHYDCT)
│   └── bieu-mau-quan-ly-noi-bo.md           # Chi tiết 10 mẫu VBQLNB (QĐ 1494/QĐ-ĐHYDCT)
└── resources/
    ├── 01-quan-ly-noi-bo-qd1494/            # 10 tệp .docx biểu mẫu VBQLNB chuẩn (VP.VBQLNB.01 → 10)
    └── 02-hanh-chinh-qd1488/                # 12 tệp .docx biểu mẫu hành chính chuẩn (Mẫu 1 → 11)
```

---

## 🛠 Hướng dẫn Cài đặt & Sử dụng

### Dành cho Google Antigravity
1. Đặt thư mục `skills/tao-bien-ban-cuoc-hop` vào `.agents/skills/tao-bien-ban-cuoc-hop` của workspace.
2. Hoặc cấu hình `.agents/skills.json` trong dự án:
```json
{
  "entries": [
    {
      "path": "skills"
    }
  ]
}
```
3. Trong giao diện chat với Antigravity, bạn chỉ cần gõ yêu cầu:
   > *"Dựa vào file ghi chú cuộc họp [duong_dan_file], hãy dùng skill `tao-bien-ban-cuoc-hop` để tạo biên bản cuộc họp chuẩn Trường Đại học Y Dược Cần Thơ."*

---

## 👤 Tác giả
- **Tác giả:** KTV. Nguyễn Phúc Duy
- **Đơn vị:** Bộ môn Mô phôi - GPB - Pháp Y, Khoa Y, Trường Đại học Y Dược Cần Thơ / NPDUYLAB
- **GitHub:** [@npduylab](https://github.com/npduylab)
- **Website:** [npduylab.com](https://npduylab.com/)

---

## 📄 Bản quyền (License)
Phát hành theo giấy phép [MIT License](LICENSE).
