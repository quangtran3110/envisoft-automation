"""
File chính - chạy toàn bộ quy trình tự động hóa:
1. Đăng nhập + tải báo cáo từ website
2. Đọc Excel + ghi vào Google Sheet

Cách chạy:
    python main.py
"""

import os
import sys
from dotenv import load_dotenv

from fetch_reports import fetch_all_reports
from write_to_sheet import (
    get_google_sheet_client,
    process_excel_file,
)


def main():
    # Load biến môi trường từ file .env (khi chạy local).
    # Khi chạy trên GitHub Actions, biến env đã có sẵn nên load_dotenv không ảnh hưởng.
    load_dotenv()

    username = os.getenv("ENVISOFT_USERNAME")
    password = os.getenv("ENVISOFT_PASSWORD")
    sheet_id = os.getenv("GOOGLE_SHEET_ID")
    target_year = int(os.getenv("TARGET_YEAR", "2026"))
    days_to_fetch = int(os.getenv("DAYS_TO_FETCH", "14"))
    # Đặt HEADLESS=false trong .env khi muốn xem trình duyệt làm việc (debug local).
    headless = os.getenv("HEADLESS", "true").lower() != "false"

    missing = []
    if not username:
        missing.append("ENVISOFT_USERNAME")
    if not password:
        missing.append("ENVISOFT_PASSWORD")
    if not sheet_id:
        missing.append("GOOGLE_SHEET_ID")

    if missing:
        print(f"Thiếu cấu hình: {', '.join(missing)}")
        print("   Kiểm tra file .env (chạy local) hoặc GitHub Secrets (CI).")
        sys.exit(1)

    print("=" * 60)
    print("BẮT ĐẦU QUY TRÌNH TỰ ĐỘNG HÓA")
    print("=" * 60)

    # ===== BƯỚC 1: TẢI BÁO CÁO =====
    print("\n[1/2] Đang tải báo cáo từ website Envisoft...")
    try:
        downloaded_files = fetch_all_reports(
            username, password, days_to_fetch, headless=headless
        )
    except Exception as e:
        print(f"Lỗi khi tải báo cáo: {e}")
        sys.exit(1)

    if not downloaded_files:
        print("Không có file nào được tải về. Kết thúc.")
        sys.exit(0)

    print(f"\nĐã tải {len(downloaded_files)} file.")

    # ===== BƯỚC 2: GHI VÀO GOOGLE SHEET =====
    print("\n[2/2] Đang ghi dữ liệu vào Google Sheet...")
    try:
        gs_client = get_google_sheet_client()
    except Exception as e:
        print(f"Lỗi kết nối Google Sheet: {e}")
        sys.exit(1)

    results = []
    for item in downloaded_files:
        excel_path = item["path"]
        well_info = item["well"]
        sheet_name = well_info["sheet_name"]

        try:
            result = process_excel_file(
                excel_path=excel_path,
                sheet_name=sheet_name,
                gspread_client=gs_client,
                sheet_id=sheet_id,
                target_year=target_year,
                force=True,
            )

            results.append({
                "file": excel_path.name,
                "sheet": sheet_name,
                "result": result,
            })
        except Exception as e:
            print(f"   Lỗi xử lý {excel_path.name}: {e}")
            results.append({
                "file": excel_path.name,
                "sheet": sheet_name,
                "result": {"status": "error", "message": str(e)},
            })

    # ===== TỔNG KẾT =====
    print("\n" + "=" * 60)
    print("KẾT QUẢ TỔNG KẾT")
    print("=" * 60)
    success_count = sum(1 for r in results if r["result"]["status"] == "success")
    print(f"Thành công: {success_count}/{len(results)} file")

    for r in results:
        if r["result"]["status"] == "success":
            print(f"  OK  {r['file']:40s} -> Sheet '{r['sheet']}' "
                  f"(+{r['result']['added']} ô)")
        else:
            msg = r["result"].get("message", "")
            print(f"  XX  {r['file']:40s} -> LỖI: {msg}")

    print("\nHOÀN THÀNH.\n")

    # Exit code khác 0 nếu có file lỗi để GitHub Actions báo failed
    if success_count < len(results):
        sys.exit(2)


if __name__ == "__main__":
    main()
