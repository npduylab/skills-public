---
name: luu-kho-git
description: >-
  Quy trình chuẩn hóa việc kiểm tra, phân loại (Public/Private), đồng bộ, commit và push skill,
  mã nguồn hoặc tài liệu lên GitHub từ Hub trung tâm d:\Antigravity_system, Antigravity global
  (.gemini/config/skills) hoặc Google Drive (Y:\) thông qua các bản clone cục bộ và GitHub CLI gh.
  Kích hoạt khi người dùng yêu cầu lưu, commit, push, đồng bộ skill lên 2 kho npduylab/skills-public
  và npduylab/skills-private, hoặc tạo/cập nhật repo GitHub khác.
---

# Lưu Kho Git & Đồng Bộ Skills Lên GitHub (`luu-kho-git`)

Skill này chuẩn hóa quy trình quản lý, phân loại Công khai/Riêng tư, đồng bộ đa điểm và đẩy mã nguồn/skill lên GitHub an toàn (không đặt `.git` trực tiếp trên ổ Google Drive `Y:\` và không tự ý xóa dữ liệu).

---

## 1. Kiến Trúc Kho Lưu Trữ & Đồng Bộ Chuẩn

- **Tài khoản GitHub:** `npduylab` (`NGUYỄN PHÚC DUY <npduylab@gmail.com>`)
- **Công cụ:** `git` + GitHub CLI (`gh`)
- **Hub điều phối trung tâm (Single Source of Truth):** `d:\Antigravity_system\`
  - `d:\Antigravity_system\skills-manifest.json` — Danh sách khai báo skill thuộc `public` hay `private`
  - `d:\Antigravity_system\public\` — Nguồn các skill công khai
  - `d:\Antigravity_system\private\` — Nguồn các skill riêng tư / nội bộ
  - `d:\Antigravity_system\scripts\sync-skills.ps1` — Script đồng bộ tự động
  - `d:\Antigravity_system\scripts\check-secrets.ps1` — Script quét rác & rò rỉ khóa bí mật
- **2 Kho GitHub chuyên biệt cho Skills (clone tại `C:\Users\Client\source\repos\`):**
  1. **Skill Công khai (`PUBLIC`):** `https://github.com/npduylab/skills-public` → `C:\Users\Client\source\repos\skills-public`
     - Dành cho skill tiện ích dùng chung, không chứa thông tin nhạy cảm/nội bộ riêng tư.
  2. **Skill Riêng tư (`PRIVATE`):** `https://github.com/npduylab/skills-private` → `C:\Users\Client\source\repos\skills-private`
     - Dành cho skill đặc thù nội bộ Bộ môn/CTUMP, quy trình đăng nhập, Google Workspace nội bộ, cấu hình cá nhân và bản sao lưu `hub/` (`skills-manifest.json`, `scripts/`).
- **2 Điểm chạy thực tế được đồng bộ từ Hub:**
  - **Antigravity Global Runtime:** `C:\Users\Client\.gemini\config\skills\`
  - **Google Drive dùng chung (`Y:\`):** `Y:\.agents\skills\`

> **Nguyên tắc bắt buộc:**
> 1. **KHÔNG** chạy `git init` trực tiếp trên ổ Google Drive (`Y:\`, `W:\`, `X:\`, `Z:\`) để tránh xung đột đồng bộ RaiDrive.
> 2. **KHÔNG** tự ý xóa tệp hoặc thư mục ở bất kỳ vị trí nào khi chưa có sự phê duyệt rõ ràng của người dùng.
> 3. **Mặc định an toàn:** Skill mới chưa khai báo trong `skills-manifest.json` luôn xếp vào `private/` trước cho đến khi người dùng xác nhận công khai.

---

## 2. Quy Trình Cập Nhật & Đồng Bộ Skills

### Trường hợp A: Kiểm tra trạng thái đồng bộ hiện tại
```powershell
powershell -ExecutionPolicy Bypass -File "d:\Antigravity_system\scripts\sync-skills.ps1" -Mode Status
```
Hiển thị bảng đối chiếu giữa `Hub`, `Gemini (.gemini)`, `YDrive (Y:\)` và `GitClone`, đồng thời cảnh báo nếu có skill mới chưa đăng ký trong `skills-manifest.json`.

### Trường hợp B: Sửa hoặc tạo skill tại Hub (`d:\Antigravity_system`) rồi đẩy lên GitHub
1. Chỉnh sửa/tạo skill trong `d:\Antigravity_system\public\<ten-skill>` hoặc `d:\Antigravity_system\private\<ten-skill>`.
2. Nếu là skill mới, khai báo vào `d:\Antigravity_system\skills-manifest.json` (và cập nhật bảng danh mục trong `README.md` của repo tương ứng).
3. Chạy đồng bộ & push (tự động deploy ra `.gemini`, `Y:\`, quét an toàn `check-secrets.ps1`, commit & push):
   ```powershell
   # Đẩy cả 2 kho hoặc chọn -Target public / -Target private
   powershell -ExecutionPolicy Bypass -File "d:\Antigravity_system\scripts\sync-skills.ps1" -Mode Push -Target all -Message "feat: cap nhat skill <ten-skill>"
   ```

### Trường hợp C: Lỡ sửa/tạo skill trực tiếp tại `.gemini\config\skills\` hoặc `Y:\.agents\skills\`
1. Thu gom bản mới từ `.gemini` (hoặc `ydrive`) về Hub:
   ```powershell
   powershell -ExecutionPolicy Bypass -File "d:\Antigravity_system\scripts\sync-skills.ps1" -Mode Collect -Source gemini
   ```
2. Kiểm tra lại phân loại trong `skills-manifest.json` (nếu là skill mới, mặc định đã vào `private/`).
3. Chạy `-Mode Push` để chuẩn hóa đồng bộ ngược lại tất cả điểm và đẩy lên GitHub.

### Trường hợp D: Đồng bộ từ GitHub về máy (máy mới hoặc có commit từ nơi khác)
```powershell
powershell -ExecutionPolicy Bypass -File "d:\Antigravity_system\scripts\sync-skills.ps1" -Mode Pull
```
Lệnh sẽ `git pull --ff-only` cả 2 repo, cập nhật vào Hub và tự động deploy ra `.gemini\config\skills\` và `Y:\.agents\skills\`.

---

## 3. Quy Trình Chung Cho Các Dự Án / Repo GitHub Khác

Khi cần đẩy mã nguồn hoặc tài liệu không thuộc bộ Skills lên một repo khác của `npduylab`:

1. **Kiểm tra đăng nhập `gh`:**
   ```powershell
   $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
   gh auth status
   ```
2. **Xác định repo đích:**
   - Liệt kê repo hiện có: `gh repo list npduylab --limit 50 --json name,url,visibility,updatedAt`
   - Nếu tạo repo mới: hỏi rõ người dùng **tên repo** và **`--public` hay `--private`**, sau đó clone về `C:\Users\Client\source\repos\<ten-repo>`.
3. **Rà soát an toàn:** Quét tệp rác (`~$*`, `__pycache__/`, `*.pyc`, `Thumbs.db`, `desktop.ini`, `.DS_Store`) và bí mật (`.env`, API key, token `gho_`, `AIza...`, `sk-...`). Với repo `PUBLIC`, xác nhận lại với người dùng trước khi `push`.
4. **Commit & Push:**
   ```powershell
   $repo = "C:\Users\Client\source\repos\<ten-repo>"
   git -C $repo config user.name "NGUYỄN PHÚC DUY"
   git -C $repo config user.email "npduylab@gmail.com"
   git -C $repo add -A
   git -C $repo status -s
   git -C $repo commit -m "<type>: <mo ta ngan gon>"
   git -C $repo push
   ```
   - Tiền tố commit chuẩn: `feat:`, `fix:`, `docs:`, `chore:`.
5. **Xác thực kết quả:** Chạy `git -C $repo log -1 --oneline` và `git -C $repo status -s`, báo rõ tên repo, chế độ `PUBLIC`/`PRIVATE`, mã commit và liên kết GitHub.
