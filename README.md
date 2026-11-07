# NPDUYLAB AI Agent Skills Collection 🚀

Kho lưu trữ các bộ kỹ năng (Skills) chuyên biệt dành cho **AI Coding Assistants & Autonomous Agents** (Google Antigravity, Claude Code, Cursor, Codex...), được phát triển và tối ưu bởi **Nguyễn Phúc Duy (NPDUYLAB)**.

---

## 📂 Danh mục Kỹ năng (Skills Catalog)

| Kỹ năng | Phiên bản | Mô tả tóm tắt | Nền tảng hỗ trợ |
| :--- | :---: | :--- | :--- |
| [`tao-bien-ban-cuoc-hop`](./skills/tao-bien-ban-cuoc-hop/SKILL.md) | `1.0.0` | Tự động chuyển đổi ghi chép thô thành **Biên bản cuộc họp chính thức** chuẩn thể thức hành chính Việt Nam theo **Nghị định 30/2020/NĐ-CP**. Hỗ trợ xuất đồng thời Markdown (.md) và Word (.docx). | Antigravity, Claude Code, Cursor, Codex |

---

## 📌 Chi tiết Kỹ năng: `tao-bien-ban-cuoc-hop`

### 1. Giới thiệu
Kỹ năng hướng dẫn AI Agent tiếp nhận các bản ghi chép nhanh, bullet points, transcript cuộc họp và biên soạn thành biên bản cuộc họp chính thức theo quy chuẩn quản lý nhà nước:
- **Quốc hiệu, Tiêu ngữ, Tên cơ quan, Số/Ký hiệu, Trích yếu** chuẩn văn bản hành chính.
- **5 Phần cốt lõi:** Thời gian địa điểm, Thành phần tham dự, Nội dung diễn biến, Kết luận chủ trì (kèm bảng phân công hành động Action Items), Chữ ký xác nhận thư ký và chủ tọa.
- **Công cụ chuyển đổi Word tự động (`export_docx.py`):** Căn lề chuẩn A4 (Trái 30mm, Phải 20mm, Trên/Dưới 20mm), font Times New Roman, kẻ viền bảng phân công, ẩn viền bảng chữ ký theo chuẩn ISO & NĐ 30/2020/NĐ-CP.

### 2. Cấu trúc thư mục kỹ năng
```text
skills/tao-bien-ban-cuoc-hop/
├── SKILL.md                          # Tài liệu hướng dẫn & quy trình 5 bước cho Agent
├── scripts/
│   └── export_docx.py                # Script Python chuyển đổi Markdown/JSON sang Word (.docx) chuẩn
├── references/
│   ├── nghi-dinh-30-2020-nd-cp.md    # Bảng tra cứu quy chuẩn thể thức NĐ 30/2020/NĐ-CP
│   └── templates.md                  # Bộ khung mẫu biên bản giao ban & họp chuyên môn
└── examples/
    ├── ghi_chu_mau.md                # File dữ liệu ghi chú mẫu đầu vào
    ├── bien_ban_chuan_mau.md         # File biên bản hoàn chỉnh mẫu đầu ra (.md)
    └── sample_output.docx            # File Word thành phẩm mẫu (.docx)
```

---

## 🛠 Hướng dẫn Cài đặt & Sử dụng

### Dành cho Google Antigravity
1. Clone hoặc tải thư mục `skills/tao-bien-ban-cuoc-hop` về dự án của bạn (đặt vào `.agents/skills/tao-bien-ban-cuoc-hop`).
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
   > *"Dựa vào file ghi chú cuộc họp [duong_dan_file], hãy dùng skill `tao-bien-ban-cuoc-hop` để tạo biên bản cuộc họp chính thức."*

### Yêu cầu môi trường cho script xuất Word:
Nếu muốn chạy script `export_docx.py` tự động:
```bash
pip install python-docx
```

---

## 👤 Tác giả
- **Tác giả:** KTV. Nguyễn Phúc Duy
- **Đơn vị:** NPDUYLAB
- **GitHub:** [@npduylab](https://github.com/npduylab)
- **Website:** [npduylab.com](https://npduylab.com/)

---

## 📄 Bản quyền (License)
Phát hành theo giấy phép [MIT License](LICENSE).
