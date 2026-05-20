"""
Module này port lại logic của code.gs (processFile) sang Python:
1. Đọc file Excel báo cáo
2. Tìm các cột Datetime, Lưu lượng, Mực Nước
3. Parse từng dòng dữ liệu, ghi vào Google Sheet theo cấu trúc:
   - Hàng bắt đầu: 4 (tháng 1)
   - Cột bắt đầu: D (cột 4)
   - 24 hàng (12 tháng × 2 hàng: Q và H)
   - 31 cột (31 ngày)
"""

import os
import openpyxl
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime


# Cấu hình giống code.gs cũ
START_ROW = 4         # Hàng bắt đầu của bảng dữ liệu (tháng 1)
START_COL = 4         # Cột D
NUM_ROWS = 24         # 12 tháng × 2 hàng (Q và H)
NUM_COLS = 31         # 31 ngày


def get_google_sheet_client():
    """
    Khởi tạo client kết nối Google Sheets.
    Ưu tiên đọc credentials từ biến môi trường (cho GitHub Actions),
    fallback sang file JSON (cho chạy local).
    """
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    # Ưu tiên 1: biến môi trường GOOGLE_SERVICE_ACCOUNT_JSON (chuỗi JSON đầy đủ)
    creds_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    if creds_json:
        import json
        creds_dict = json.loads(creds_json)
        credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    else:
        # Ưu tiên 2: file JSON
        creds_path = os.getenv("GOOGLE_CREDENTIALS_PATH", "./service_account.json")
        if not os.path.exists(creds_path):
            raise FileNotFoundError(
                f"Không tìm thấy file credentials: {creds_path}\n"
                "Hãy đặt file service_account.json vào thư mục dự án, "
                "hoặc set biến môi trường GOOGLE_SERVICE_ACCOUNT_JSON."
            )
        credentials = Credentials.from_service_account_file(creds_path, scopes=scopes)

    return gspread.authorize(credentials)


# NOTE: hàm extract_well_from_filename (port từ extractWellFromFileName trong code.gs)
# đã được lược bỏ. Trong bản tự động, sheet_name được lấy trực tiếp từ wells.txt
# (cột thứ 2: "MH113a|113a") chứ không phụ thuộc vào tên file Excel.


def parse_datetime_cell(value):
    """
    Parse cell datetime trong Excel, trả về (day, month, year).
    Hỗ trợ cả datetime object và string.
    """
    if value is None or value == "":
        return None

    # Nếu là datetime object (Excel parse tự động)
    if isinstance(value, datetime):
        return {"day": value.day, "month": value.month, "year": value.year}

    # Nếu là string, parse như code.gs cũ
    try:
        s = str(value).strip()
        parts = s.split(" ")
        date_part = parts[1] if len(parts) > 1 else parts[0]
        day, month, year = map(int, date_part.split("/"))
        return {"day": day, "month": month, "year": year}
    except Exception:
        return None


def find_column_index(headers, keyword):
    """Tìm index của cột chứa keyword (case-insensitive)."""
    for i, h in enumerate(headers):
        if h and keyword.lower() in str(h).lower():
            return i
    return -1


def process_excel_file(excel_path, sheet_name, gspread_client, sheet_id, target_year, force=True):
    """
    Đọc file Excel, parse dữ liệu, ghi vào sheet tương ứng trên Google Sheet.

    excel_path:    đường dẫn file Excel
    sheet_name:    tên sheet đích trong Google Sheet (ví dụ "1", "3a")
    sheet_id:      ID của Google Spreadsheet
    target_year:   năm cần xử lý (chỉ ghi data có year khớp)
    force:         True = ghi đè dữ liệu cũ (mặc định khi tự động hóa)
    """
    print(f"   → Xử lý file: {excel_path}")

    # 1. Đọc Excel
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    if len(rows) < 2:
        return {"status": "error", "message": "File không có dữ liệu"}

    # 2. Tìm các cột cần thiết
    headers = [str(h).strip() if h else "" for h in rows[0]]
    datetime_col = -1
    for i, h in enumerate(headers):
        if h == "Datetime":
            datetime_col = i
            break

    q_col = find_column_index(headers, "Lưu lượng")
    h_col = find_column_index(headers, "Mực Nước")

    if datetime_col == -1 or q_col == -1 or h_col == -1:
        return {
            "status": "error",
            "message": f"Không tìm thấy cột cần thiết (Datetime/Lưu lượng/Mực Nước). "
                      f"Headers: {headers}"
        }

    # 3. Mở Google Sheet
    spreadsheet = gspread_client.open_by_key(sheet_id)
    try:
        worksheet = spreadsheet.worksheet(sheet_name)
    except gspread.WorksheetNotFound:
        return {"status": "error", "message": f"Không tìm thấy sheet '{sheet_name}'"}

    # 4. Đọc block dữ liệu hiện tại 1 lần (giống code.gs)
    # Range: D4:AH27 (24 hàng × 31 cột bắt đầu từ D4)
    range_str = f"{gspread.utils.rowcol_to_a1(START_ROW, START_COL)}:" \
                f"{gspread.utils.rowcol_to_a1(START_ROW + NUM_ROWS - 1, START_COL + NUM_COLS - 1)}"

    block_values = worksheet.get(range_str)

    # Chuẩn hóa block về đúng kích thước 24×31 (gspread bỏ qua hàng/cột trống ở cuối)
    block = []
    for r in range(NUM_ROWS):
        row = block_values[r] if r < len(block_values) else []
        # Pad cho đủ 31 cột
        row = list(row) + [""] * (NUM_COLS - len(row))
        block.append(row[:NUM_COLS])

    # 5. Xử lý từng dòng dữ liệu trong Excel
    duplicates = []
    added_count = 0

    for row_data in rows[1:]:
        if not row_data or len(row_data) <= max(datetime_col, q_col, h_col):
            continue

        parsed = parse_datetime_cell(row_data[datetime_col])
        if not parsed:
            continue
        if parsed["year"] != target_year:
            continue

        day = parsed["day"]
        month = parsed["month"]
        if not (1 <= day <= 31 and 1 <= month <= 12):
            continue

        row_index_q = (month - 1) * 2
        row_index_h = row_index_q + 1
        col_index = day - 1

        existing = block[row_index_q][col_index]
        if existing not in ("", None) and not force:
            duplicates.append(f"{day}/{month}")
            continue

        # Convert sang float để Google Sheets nhận là NUMBER (không phải TEXT).
        # Nếu giá trị không parse được (None, "", N/A...) thì để nguyên để fall qua.
        def _to_num(v):
            if v is None or v == "":
                return ""
            try:
                return float(v)
            except (ValueError, TypeError):
                return v

        block[row_index_q][col_index] = _to_num(row_data[q_col])
        block[row_index_h][col_index] = _to_num(row_data[h_col])
        added_count += 1

    # 6. Ghi lại block 1 lần
    # value_input_option='USER_ENTERED' = Google Sheets tự parse như khi gõ tay,
    # nên float 43.8081 sẽ là NUMBER, không phải '43.8081 TEXT.
    if added_count > 0:
        worksheet.update(range_name=range_str, values=block,
                         value_input_option='USER_ENTERED')

    print(f"      ✓ Sheet '{sheet_name}': +{added_count} ô, {len(duplicates)} ngày trùng")

    return {
        "status": "success",
        "well": sheet_name,
        "added": added_count,
        "duplicates": len(duplicates)
    }


def log_to_sheet(gspread_client, sheet_id, file_name, well, rows):
    """Ghi log vào sheet LOG_UPLOAD (giống logUpload trong code.gs)."""
    try:
        spreadsheet = gspread_client.open_by_key(sheet_id)
        try:
            log_ws = spreadsheet.worksheet("LOG_UPLOAD")
        except gspread.WorksheetNotFound:
            log_ws = spreadsheet.add_worksheet(title="LOG_UPLOAD", rows=1000, cols=4)
            log_ws.append_row(["Thời gian", "File", "Giếng", "Số ngày cập nhật"])

        log_ws.append_row([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            file_name,
            well,
            rows
        ])
    except Exception as e:
        print(f"   (Cảnh báo: không ghi được log: {e})")
