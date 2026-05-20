"""
Module tải báo cáo Excel từ website la.envisoft.gov.vn:8090

Đã verify selector trực tiếp trên trang thật (qua Claude in Chrome).

Cấu trúc trang:
- Login form (web2py, có _formkey CSRF tự động)
- Trang Tra cứu (/eos/view_log/average_data_by_time) dùng Chosen.js cho dropdown
  Tuy nhiên ta KHÔNG dùng UI Chosen — set thẳng vào <select> gốc + trigger jQuery
- Nút "Xuất dữ liệu Excel" (#btnExportExcel) dùng window.open() để tải file

Selectors xác nhận:
  Login:        input[name="username"], input[name="password"], button:has-text("Đăng nhập")
  View type:    select[name="view_type"]  — set "Trung bình theo ngày"
  Env comp:     select[name="station_type"]  — set value="2" (Nước ngầm NM)
  Province:     select[name="province_id"]  — mặc định Long An
  Station:      select[name="station_id"]  — set value theo option khớp keyword
  Date from:    input[name="from_date"]  — định dạng YYYY-MM-DD
  Date to:      input[name="to_date"]  — định dạng YYYY-MM-DD
  Search btn:   button.btnCustomSearch  — text "Tìm kiếm"
  Export btn:   button#btnExportExcel  — text "Xuất dữ liệu Excel"
"""

import os
from datetime import datetime, timedelta
from pathlib import Path
from playwright.sync_api import sync_playwright


BASE_URL = "http://la.envisoft.gov.vn:8090"
LOGIN_URL = f"{BASE_URL}/eos/master/login"
DATA_LOOKUP_URL = f"{BASE_URL}/eos/view_log/average_data_by_time"
DOWNLOAD_DIR = Path("downloads")

# Giới hạn từ website: xuất tối đa 1 tháng (~30 ngày)/lần
MAX_DAYS_PER_EXPORT = 30
DEFAULT_TIMEOUT = 20000


def read_wells_list(wells_file="wells.txt"):
    """Đọc wells.txt. Mỗi dòng: <search_keyword>|<tên_sheet>.

    search_keyword là chuỗi (case-insensitive) để TÌM trong text của <option>
    trong dropdown trạm. Ví dụ "MH113a" sẽ khớp:
        "CN Kiến Tường - Nhà máy nước số 2 - Giếng MH113a (NN)"
    """
    wells = []
    with open(wells_file, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            parts = line.split("|")
            if len(parts) == 2:
                wells.append({
                    "search_keyword": parts[0].strip(),
                    "sheet_name": parts[1].strip(),
                })
    return wells


def login_to_envisoft(page, username, password):
    """Đăng nhập web2py. Playwright tự gửi kèm _formkey/_formname khi submit."""
    print(f"   -> Mở trang login")
    page.goto(LOGIN_URL, wait_until="networkidle", timeout=DEFAULT_TIMEOUT)
    page.fill('input[name="username"]', username)
    page.fill('input[name="password"]', password)
    page.click('button:has-text("Đăng nhập"), input[type="submit"]')
    page.wait_for_load_state("networkidle", timeout=DEFAULT_TIMEOUT)
    if "/login" in page.url.lower():
        raise Exception(
            "Đăng nhập thất bại. Kiểm tra username/password trong .env / GitHub Secrets."
        )
    print(f"   OK đăng nhập")


def _set_select_via_jquery(page, name, value=None, label=None, regex=None):
    """Bypass Chosen.js: set value trực tiếp vào <select> gốc rồi trigger event.

    Đúng 1 trong 3 tham số được dùng:
      value: khớp option.value đúng tuyệt đối
      label: khớp option.text đúng tuyệt đối (trim)
      regex: khớp option.text bằng RegExp (case-insensitive)

    Trả về text của option được chọn, hoặc raise nếu không tìm thấy.
    """
    js = """
    ({name, value, label, regex}) => {
        const sel = document.querySelector(`select[name="${name}"]`);
        if (!sel) return {error: 'select not found: ' + name};
        let opt = null;
        if (value !== null && value !== undefined) {
            opt = Array.from(sel.options).find(o => o.value === value);
        } else if (label !== null && label !== undefined) {
            opt = Array.from(sel.options).find(o => o.text.trim() === label);
        } else if (regex !== null && regex !== undefined) {
            const re = new RegExp(regex, 'i');
            opt = Array.from(sel.options).find(o => re.test(o.text));
        }
        if (!opt) return {error: 'option not found'};
        sel.value = opt.value;
        if (window.jQuery) {
            window.jQuery(sel).trigger('change').trigger('chosen:updated');
        } else {
            sel.dispatchEvent(new Event('change', {bubbles: true}));
        }
        return {ok: true, text: opt.text, value: opt.value};
    }
    """
    result = page.evaluate(js, {
        "name": name,
        "value": value,
        "label": label,
        "regex": regex,
    })
    if result.get("error"):
        raise Exception(
            f"Không set được select '{name}' (value={value}, label={label}, regex={regex}): "
            f"{result['error']}"
        )
    return result["text"]


def setup_lookup_form(page):
    """Vào trang Tra cứu, đặt 2 dropdown CỐ ĐỊNH (1 lần cho mọi giếng)."""
    print(f"   -> Vào trang Tra cứu")
    page.goto(DATA_LOOKUP_URL, wait_until="networkidle", timeout=DEFAULT_TIMEOUT)
    page.wait_for_timeout(1500)

    # 1) Đơn vị thời gian: Trung bình theo ngày
    txt = _set_select_via_jquery(
        page, name="view_type", label="Trung bình theo ngày"
    )
    print(f"      view_type = '{txt}'")
    page.wait_for_timeout(500)

    # 2) Thành phần môi trường: Nước ngầm (NM) — value=2
    txt = _set_select_via_jquery(
        page, name="station_type", value="2"
    )
    print(f"      station_type = '{txt}'")
    # Đợi AJAX filter lại dropdown station theo station_type
    page.wait_for_timeout(2000)


def _select_all_added_columns(page):
    """Chọn TẤT CẢ option trong cbbAddedColumns (multi-select).

    Sau khi đổi station_id, website auto-populate các thông số có sẵn cho
    trạm đó (Lưu lượng, Mực Nước...). Cần đảm bảo MỌI thông số đều được
    chọn trước khi bấm Tìm kiếm — nếu không Excel xuất ra sẽ thiếu cột.
    """
    js = """
    () => {
        const sel = document.querySelector('select[name="added_columns"]');
        if (!sel) return {error: 'no added_columns select'};
        Array.from(sel.options).forEach(o => o.selected = true);
        if (window.jQuery) {
            window.jQuery(sel).trigger('change').trigger('chosen:updated');
        } else {
            sel.dispatchEvent(new Event('change', {bubbles: true}));
        }
        return {
            ok: true,
            options: Array.from(sel.options).map(o => o.text),
        };
    }
    """
    result = page.evaluate(js)
    if result.get("error"):
        raise Exception(f"Không chọn được added_columns: {result['error']}")
    return result["options"]


def fetch_well_report(page, well, date_from, date_to, output_path):
    """Chọn 1 giếng, đặt ngày, tìm kiếm, tải Excel."""
    keyword = well["search_keyword"]
    print(f"\n   -> Giếng: {keyword}")

    # 1) Chọn trạm theo keyword (case-insensitive substring trên text)
    try:
        station_text = _set_select_via_jquery(
            page, name="station_id", regex=keyword
        )
        print(f"      Trạm: {station_text}")
    except Exception as e:
        raise Exception(f"Không khớp được trạm với '{keyword}': {e}")
    # Đợi AJAX populate cbbAddedColumns theo các thông số có sẵn của trạm
    page.wait_for_timeout(1500)

    # 1b) Chọn TẤT CẢ thông số (Lưu lượng + Mực Nước) trước khi search.
    # Nếu bỏ qua bước này, Excel sẽ chỉ có cột Ngày giờ, không có data!
    try:
        cols = _select_all_added_columns(page)
        print(f"      Thông số: {', '.join(cols)}")
    except Exception as e:
        print(f"      (cảnh báo) không set được added_columns: {e}")

    # 2) Đặt 2 ô ngày
    date_from_str = date_from.strftime("%Y-%m-%d")
    date_to_str = date_to.strftime("%Y-%m-%d")
    print(f"      Ngày: {date_from_str} -> {date_to_str}")
    page.fill('input[name="from_date"]', date_from_str)
    page.fill('input[name="to_date"]', date_to_str)
    page.keyboard.press("Escape")  # đóng datepicker popup nếu có
    page.wait_for_timeout(300)

    # 3) Bấm Tìm kiếm
    print("      Bấm Tìm kiếm")
    page.click("button.btnCustomSearch")
    page.wait_for_load_state("networkidle", timeout=DEFAULT_TIMEOUT)
    page.wait_for_timeout(1500)  # buffer cho bảng kết quả render

    # 4) Bấm Xuất dữ liệu Excel và bắt download
    # Lưu ý: nút dùng window.open() để mở URL export — Playwright cần
    # accept_downloads=True + expect_download để bắt được file
    print("      Bấm Xuất dữ liệu Excel")
    with page.expect_download(timeout=45000) as dl_info:
        page.click("button#btnExportExcel")
    download = dl_info.value
    download.save_as(output_path)
    print(f"      OK lưu: {output_path}")


def fetch_all_reports(username, password, days_to_fetch=14, headless=True):
    """Đăng nhập, lặp qua mọi giếng trong wells.txt, tải Excel cho từng giếng.

    days_to_fetch: số ngày lùi lại từ HÔM QUA (giới hạn website là 30).
                   Mặc định 14.
    headless:      True (GitHub Actions), False để debug local.
    """
    DOWNLOAD_DIR.mkdir(exist_ok=True)

    if days_to_fetch > MAX_DAYS_PER_EXPORT:
        print(f"(cảnh báo) days_to_fetch={days_to_fetch} vượt giới hạn website "
              f"({MAX_DAYS_PER_EXPORT}). Tự động cắt xuống {MAX_DAYS_PER_EXPORT}.")
        days_to_fetch = MAX_DAYS_PER_EXPORT

    yesterday = datetime.now() - timedelta(days=1)
    date_to = yesterday
    date_from = yesterday - timedelta(days=days_to_fetch - 1)
    print(f"\nKhoảng thời gian: "
          f"{date_from.strftime('%Y-%m-%d')} -> {date_to.strftime('%Y-%m-%d')} "
          f"({days_to_fetch} ngày)")

    wells = read_wells_list()
    print(f"Số giếng: {len(wells)}")

    downloaded_files = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context(
            accept_downloads=True,
            ignore_https_errors=True,
        )
        page = context.new_page()
        page.set_default_timeout(DEFAULT_TIMEOUT)
        try:
            login_to_envisoft(page, username, password)
            setup_lookup_form(page)
            for well in wells:
                safe_name = well["search_keyword"].replace("/", "_").replace(" ", "_")
                output_path = DOWNLOAD_DIR / (
                    f"{safe_name}_{date_to.strftime('%Y%m%d')}.xlsx"
                )
                try:
                    fetch_well_report(page, well, date_from, date_to, output_path)
                    downloaded_files.append({"path": output_path, "well": well})
                except Exception as e:
                    print(f"      XX LỖI giếng '{well['search_keyword']}': {e}")
                    # Reload trang Tra cứu để reset state cho giếng tiếp theo
                    try:
                        setup_lookup_form(page)
                    except Exception:
                        pass
        finally:
            browser.close()
    return downloaded_files


if __name__ == "__main__":
    from dotenv import load_dotenv
    load_dotenv()
    u = os.getenv("ENVISOFT_USERNAME")
    pw = os.getenv("ENVISOFT_PASSWORD")
    d = int(os.getenv("DAYS_TO_FETCH", "14"))
    if not u or not pw:
        print("Chưa cấu hình ENVISOFT_USERNAME/ENVISOFT_PASSWORD trong .env")
        exit(1)
    files = fetch_all_reports(u, pw, d, headless=False)
    print(f"\nTải xong {len(files)} file.")
