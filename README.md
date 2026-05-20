# HƯỚNG DẪN TỰ ĐỘNG HÓA QUY TRÌNH ENVISOFT

Tài liệu này dành cho Quang triển khai dự án từ đầu.

## 📋 TỔNG QUAN

Dự án này tự động hóa quy trình:
1. Đăng nhập website `la.envisoft.gov.vn:8090`
2. Tải báo cáo Excel cho nhiều giếng
3. Ghi dữ liệu vào Google Sheet
4. Chạy tự động hàng tuần qua GitHub Actions

---

## 🗂️ CẤU TRÚC FILE

```
envisoft-automation/
├── main.py                          ← Chạy file này
├── fetch_reports.py                 ← Module tải báo cáo
├── write_to_sheet.py                ← Module ghi Google Sheet
├── wells.txt                        ← Danh sách giếng (Quang chỉnh)
├── requirements.txt                 ← Thư viện Python cần cài
├── .env.example                     ← Mẫu cấu hình (copy thành .env)
├── .gitignore                       ← Bảo vệ file nhạy cảm
└── .github/workflows/
    └── weekly-sync.yml              ← Cấu hình chạy tự động
```

---

## 🚀 TRIỂN KHAI - 4 GIAI ĐOẠN

### GIAI ĐOẠN 1: TẠO GOOGLE SERVICE ACCOUNT (~15 phút)

Service account là "tài khoản máy" giúp Python ghi vào Google Sheet không cần OAuth phức tạp.

**Bước 1.1:** Vào https://console.cloud.google.com

**Bước 1.2:** Tạo project mới
- Click vào dropdown project ở góc trên cùng
- Click "NEW PROJECT"
- Đặt tên: `envisoft-automation` (hoặc tên Quang muốn)
- Click "CREATE"

**Bước 1.3:** Bật Google Sheets API và Drive API
- Vào menu trái → "APIs & Services" → "Library"
- Tìm "Google Sheets API" → Click → "ENABLE"
- Quay lại Library, tìm "Google Drive API" → Click → "ENABLE"

**Bước 1.4:** Tạo Service Account
- Menu trái → "IAM & Admin" → "Service Accounts"
- Click "CREATE SERVICE ACCOUNT"
- Service account name: `envisoft-bot`
- Click "CREATE AND CONTINUE"
- Role: chọn "Editor" (đơn giản nhất, hoặc "Basic > Editor")
- Click "CONTINUE" → "DONE"

**Bước 1.5:** Tạo và tải file JSON key
- Click vào service account vừa tạo
- Tab "KEYS" → "ADD KEY" → "Create new key"
- Chọn JSON → "CREATE"
- File JSON sẽ tự tải về máy. **Đổi tên thành `service_account.json`** và đặt vào thư mục dự án.

**Bước 1.6:** Share Google Sheet với service account
- Mở file `service_account.json`, tìm field `"client_email"` (dạng `xxx@xxx.iam.gserviceaccount.com`)
- Copy email đó
- Mở Google Sheet của Quang → bấm "Share" → dán email vào → đặt quyền "Editor" → "Send"

⚠️ **Lưu ý:** File `service_account.json` chứa key bí mật, **không bao giờ chia sẻ hoặc commit lên GitHub công khai**.

---

### GIAI ĐOẠN 2: CHẠY THỬ TRÊN MÁY CÁ NHÂN (~30 phút)

#### Bước 2.1: Cài đặt môi trường Python

Mở **Terminal** (Mac/Linux) hoặc **PowerShell/CMD** (Windows), `cd` vào thư mục dự án.

**Kiểm tra Python đã cài chưa:**
```bash
python --version
```
Nếu báo lỗi, thử `python3 --version`. Cần Python 3.10+.

**Tạo môi trường ảo (virtual environment):**

Đây là "hộp cách ly" để cài thư viện cho riêng dự án này, không ảnh hưởng các dự án khác.

```bash
# Windows:
python -m venv venv
venv\Scripts\activate

# Mac/Linux:
python3 -m venv venv
source venv/bin/activate
```

Sau khi kích hoạt, dòng terminal sẽ có chữ `(venv)` ở đầu.

**Cài thư viện:**
```bash
pip install -r requirements.txt
playwright install chromium
```

Lệnh thứ 2 sẽ tải Chrome cho Playwright (~150MB), chờ vài phút.

#### Bước 2.2: Cấu hình .env

Copy file mẫu:
```bash
# Windows:
copy .env.example .env

# Mac/Linux:
cp .env.example .env
```

Mở file `.env` bằng Notepad/VSCode, điền:
- `ENVISOFT_USERNAME` = tài khoản website
- `ENVISOFT_PASSWORD` = mật khẩu website
- `GOOGLE_SHEET_ID` = ID của Google Sheet (lấy từ URL, đoạn giữa `/d/` và `/edit`)

#### Bước 2.3: Cập nhật danh sách giếng

Mở file `wells.txt`, sửa danh sách giếng. Format mỗi dòng:
```
<tên trên website>|<tên sheet trong Google Sheet>
```

Ví dụ nếu trên dropdown website hiện "MH1", còn trong Google Sheet có sheet tên "1":
```
MH1|1
MH2|2
MH3a|3a
```

#### Bước 2.4: Chạy thử

**Trước khi chạy thật, mở file `fetch_reports.py` và đổi:**
```python
browser = p.chromium.launch(headless=True)
```
thành:
```python
browser = p.chromium.launch(headless=False)
```

Điều này sẽ mở trình duyệt thật để Quang **xem script làm gì** — rất quan trọng để debug lần đầu.

**Chạy script:**
```bash
python main.py
```

#### Bước 2.5: Xử lý lỗi thường gặp

⚠️ **Khả năng cao là sẽ lỗi ngay lần chạy đầu.** Đây là điều bình thường vì các selector trong code mình viết là **dự đoán**, cần điều chỉnh cho khớp với website thật. Các lỗi phổ biến:

**Lỗi: "Timeout waiting for selector"**
- Nguyên nhân: tên field/dropdown trong code không khớp với website
- Cách sửa: Mở DevTools (F12) trên website, inspect phần tử mà script đang tìm, xem `name`/`id` thật là gì → sửa trong `fetch_reports.py`
- Các vị trí cần kiểm tra (mình đánh dấu `TODO: VERIFY`):
  - Tên input username/password
  - Tên dropdown chọn trạm
  - Tên input ngày
  - Selector nút "Xuất Excel"

**Lỗi: "Login failed"**
- Kiểm tra lại username/password trong `.env`
- Thử đăng nhập tay xem có thật sự vào được không

**Lỗi: "Sheet not found"**
- Kiểm tra tên sheet trong `wells.txt` có khớp với tên sheet trong Google Sheet không
- Kiểm tra đã share Google Sheet với service account chưa

**Lỗi: "Không tìm thấy cột Datetime/Lưu lượng/Mực Nước"**
- Mở file Excel mà script tải về, xem header thật là gì
- Sửa keyword trong `write_to_sheet.py` cho khớp

---

### GIAI ĐOẠN 3: ĐẨY CODE LÊN GITHUB (~15 phút)

#### Bước 3.1: Tạo tài khoản GitHub (nếu chưa có)

Vào https://github.com/signup. Tài khoản miễn phí là đủ.

#### Bước 3.2: Tạo repository PRIVATE

⚠️ **Quan trọng: chọn PRIVATE**, không phải Public, để code không bị lộ ra ngoài.

- Vào https://github.com/new
- Repository name: `envisoft-automation`
- Visibility: **Private**
- KHÔNG tick "Add README" (vì mình đã tạo rồi)
- Click "Create repository"

#### Bước 3.3: Cài Git và push code

**Cài Git nếu chưa có:** https://git-scm.com/downloads

Trong terminal, từ thư mục dự án:

```bash
git init
git add .
git commit -m "Initial automation setup"
git branch -M main
git remote add origin https://github.com/<TÊN-CỦA-QUANG>/envisoft-automation.git
git push -u origin main
```

Thay `<TÊN-CỦA-QUANG>` bằng username GitHub.

⚠️ **Trước khi push, kiểm tra:**
```bash
git status
```
Đảm bảo file `.env` và `service_account.json` **KHÔNG xuất hiện** trong danh sách "to be committed". File `.gitignore` đã chặn chúng nhưng kiểm tra lại cho chắc.

#### Bước 3.4: Thêm GitHub Secrets

Trong repository trên GitHub:
- Vào tab **Settings** → menu trái **Secrets and variables** → **Actions**
- Click **New repository secret**, lần lượt thêm 4 secrets:

| Tên secret | Giá trị |
|---|---|
| `ENVISOFT_USERNAME` | Tài khoản website |
| `ENVISOFT_PASSWORD` | Mật khẩu website |
| `GOOGLE_SHEET_ID` | ID Google Sheet |
| `GOOGLE_SERVICE_ACCOUNT_JSON` | **Toàn bộ nội dung** file `service_account.json` (copy paste nguyên cả `{...}`) |

---

### GIAI ĐOẠN 4: KÍCH HOẠT TỰ ĐỘNG (~5 phút)

#### Bước 4.1: Chạy thử trên GitHub Actions

- Vào tab **Actions** trên GitHub
- Chọn workflow "Weekly Envisoft Data Sync" ở menu trái
- Click "Run workflow" → "Run workflow" để chạy tay 1 lần

Đợi vài phút, xem log để kiểm tra. Nếu lỗi, kiểm tra log để biết bước nào sai.

#### Bước 4.2: Bật chạy tự động hàng tuần

Workflow đã được cấu hình chạy lúc **8h sáng thứ Hai hàng tuần** (cron `0 1 * * 1` UTC = 8h ICT). Không cần làm gì thêm, nó sẽ tự chạy.

Muốn đổi lịch chạy, sửa file `.github/workflows/weekly-sync.yml`, dòng `cron:`.

---

## 🔧 BẢO TRÌ

### Khi nào cần can thiệp?

- **Website đổi giao diện** → cập nhật selector trong `fetch_reports.py`
- **Đổi mật khẩu website** → cập nhật secret `ENVISOFT_PASSWORD` trên GitHub
- **Thêm giếng mới** → sửa `wells.txt` rồi commit + push

### Kiểm tra script có chạy đúng không?

- Vào tab **Actions** trên GitHub, xem các lần chạy gần nhất
- Mở sheet `LOG_UPLOAD` trong Google Sheet, xem có dòng mới mỗi tuần không

### Khi script lỗi, làm gì?

1. Vào Actions → click vào run bị lỗi → xem log
2. Tab "Artifacts" có file Excel mà script đã tải về → download xem có vấn đề gì
3. Sửa code → commit → push → chạy lại bằng "Run workflow"

---

## ❓ FAQ

**Q: Mỗi tuần tốn bao nhiêu phút GitHub Actions?**
A: ~5-10 phút/lần chạy. GitHub cho 2000 phút/tháng miễn phí cho repo private → dư rất nhiều.

**Q: Có thể tắt tự động và chỉ chạy tay không?**
A: Có. Comment đoạn `schedule:` trong workflow file. Vẫn chạy được bằng "Run workflow" thủ công.

**Q: Code có cần Apps Script cũ không?**
A: Không. Phương án mới thay thế hoàn toàn Apps Script. Quang có thể giữ Apps Script như backup (chạy tay khi cần) hoặc xóa.

**Q: Nếu một ngày website đổi sang HTTPS hoặc đổi domain thì sao?**
A: Sửa biến `LOGIN_URL` và `DATA_LOOKUP_URL` trong `fetch_reports.py`.

---

## 📞 KHI CẦN HỖ TRỢ

Khi gặp lỗi, copy toàn bộ log lỗi và gửi lại cho Claude để được trợ giúp. Đặc biệt là lỗi liên quan đến selector — đây là phần khả năng cao cần điều chỉnh sau khi chạy thử lần đầu.
