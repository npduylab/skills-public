---
name: doc-van-ban-markitdown
description: >-
  Sử dụng tiện ích Microsoft MarkItDown (markitdown[all]) đã cài đặt toàn cục để đọc và chuyển đổi
  các tệp văn bản (PDF, Word .docx, Excel .xlsx/.xls, PowerPoint .pptx, HTML, CSV, JSON, XML,
  EPUB, ZIP, ảnh/âm thanh) sang định dạng Markdown. Kích hoạt khi cần đọc nội dung tài liệu văn phòng
  hoặc trích xuất văn bản trên máy.
---

# Hướng Dẫn Sử Dụng Microsoft MarkItDown (Global)

Tiện ích [`microsoft/markitdown`](https://github.com/microsoft/markitdown) đã được cài đặt toàn cục với đầy đủ phần mở rộng (`markitdown[all]`, phiên bản `0.1.9`) trên môi trường Python 3.12.

- **Lệnh CLI:** `markitdown` (hoặc `python -m markitdown`)
- **Đường dẫn thực thi:** `C:\Users\Client\AppData\Local\Programs\Python\Python312\Scripts\markitdown.exe`

---

## 1. Lưu Ý Bắt Buộc Về Tiếng Việt (UTF-8) Trên Windows

Shell PowerShell trên Windows mặc định dùng mã hóa `cp1252` cho `stdout`, sẽ làm lỗi dấu tiếng Việt nếu chạy trực tiếp `markitdown <file>` ra màn hình. **Luôn dùng một trong hai cách sau:**

### Cách 1: Xuất ra tệp Markdown (`-o`) rồi đọc tệp (Khuyên dùng cho tệp dài)
```powershell
markitdown "duong_dan\tai_lieu.docx" -o "duong_dan_tam\output.md"
```
Sau đó đọc tệp `output.md` bằng `read_file` hoặc `Get-Content ... -Encoding utf8`.

### Cách 2: Bật UTF-8 trước khi in ra `stdout`
```powershell
$env:PYTHONUTF8 = "1"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
markitdown "duong_dan\tai_lieu.pdf"
```

---

## 2. Cách Dùng Qua Command Line (CLI)

- **Chuyển đổi cơ bản:**
  ```powershell
  $env:PYTHONUTF8 = "1"; [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
  markitdown "C:\path\to\document.docx"
  ```
- **Lưu ra tệp `.md`:**
  ```powershell
  markitdown "C:\path\to\report.xlsx" -o "C:\path\to\report.md"
  ```
- **Kiểm tra phiên bản / trợ giúp:**
  ```powershell
  markitdown --version
  markitdown --help
  ```

---

## 3. Cách Dùng Qua Python API

```python
import sys
from markitdown import MarkItDown

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

md = MarkItDown()
result = md.convert(r"C:\path\to\document.pdf")
print(result.text_content)
```

---

## 4. Định Dạng Hỗ Trợ & Lưu Ý

- **Hỗ trợ tốt:** `.docx`, `.xlsx`, `.xls`, `.pptx`, `.pdf` (có lớp văn bản), `.html`, `.htm`, `.csv`, `.json`, `.xml`, `.epub`, `.zip` (duyệt các tệp bên trong).
- **Tệp PDF dạng scan / ảnh chụp:** `markitdown` offline dùng `pdfminer.six` / `pdfplumber` nên chỉ đọc được PDF có lớp text. Nếu gặp PDF scan (trích xuất ra rỗng), hãy dùng công cụ `read_file` trực tiếp lên PDF hoặc OCR/LLM đa phương thức.
- **Bảo vệ dữ liệu:** Khi xuất tệp `.md` tạm để đọc, lưu vào thư mục scratch của phiên làm việc (`<appDataDir>\brain\<conversation-id>\scratch\`), không tạo tệp rác vào thư mục tài liệu gốc và **không tự ý xóa tệp** khi chưa được người dùng cho phép.
