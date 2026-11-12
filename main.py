import os
import sys
import time
import json
import re
import shutil
from pathlib import Path
import openpyxl
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.json"
USER_DATA_DIR = BASE_DIR / "browser_profile"


def load_config():
    if not CONFIG_PATH.exists():
        print(f"[!] Không tìm thấy file cấu hình {CONFIG_PATH}")
        sys.exit(1)
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def save_config(config):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=2, ensure_ascii=False)


def backup_excel(file_path):
    p = Path(file_path)
    if p.exists():
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        backup_path = p.with_name(f"{p.stem}_backup_{timestamp}{p.suffix}")
        try:
            shutil.copy2(p, backup_path)
            print(f"[+] Đã tạo bản sao lưu: {backup_path.name}")
        except Exception as e:
            print(f"[!] Lỗi backup: {e}")


def choose_excel_file(config):
    configured_path = config.get("excel_path", "sample_input.xlsx")
    target_path = Path(configured_path)
    if not target_path.is_absolute():
        target_path = BASE_DIR / target_path

    if target_path.exists():
        return target_path

    xlsx_files = [f for f in BASE_DIR.glob("*.xlsx") if not f.name.startswith("~$") and "backup" not in f.name]
    if "sample_input.xlsx" in [f.name for f in xlsx_files]:
        sample = BASE_DIR / "sample_input.xlsx"
        config["excel_path"] = sample.name
        save_config(config)
        return sample
    elif len(xlsx_files) >= 1:
        chosen = xlsx_files[0]
        config["excel_path"] = chosen.name
        save_config(config)
        return chosen

    return None


def parse_product_name_from_drawer(raw_text):
    """
    Trích xuất từ sau dấu '-' cho đến trước chữ 'Đã duyệt'
    Ví dụ: 'UPC: 745854124130 - KEM CHỐNG NẮNG DƯỠNG ẨM 2IN1 CETAPHIL DAILY DEFENSE CREAM SPF 50 Đã duyệt'
           -> 'KEM CHỐNG NẮNG DƯỠNG ẨM 2IN1 CETAPHIL DAILY DEFENSE CREAM SPF 50'
    """
    if not raw_text:
        return None

    # 1. Regex chính xác
    match = re.search(r"-\s*(.*?)\s*(?:Đã\s*duyệt|đã\s*duyệt|ĐÃ\s*DUYỆT)", raw_text, re.DOTALL | re.IGNORECASE)
    if match:
        name = match.group(1).strip()
        name = " ".join(name.split())
        if len(name) > 1:
            return name

    # 2. Fallback tách chuỗi
    text = raw_text
    if "-" in text:
        text = text.split("-", 1)[1]

    for marker in ["Đã duyệt", "đã duyệt", "ĐÃ DUYỆT", "Da duyet", "da duyet"]:
        idx = text.lower().find(marker.lower())
        if idx != -1:
            text = text[:idx]
            break

    name = " ".join(text.split()).strip()
    return name if len(name) > 1 else None


def wait_for_user_login(page):
    is_login_page = page.locator("input[name*='LoginInput'], button:has-text('Đăng nhập')").first.is_visible(timeout=3000)
    if is_login_page:
        print("\n" + "=" * 65)
        print("🔐 PHÁT HIỆN TRANG ĐĂNG NHẬP PIM (FPT RETAIL)")
        print("   1. Bạn hãy chuyển sang cửa sổ Chrome vừa mở.")
        print("   2. Nhập Tên đăng nhập, Mật khẩu và mã OTP 6 số.")
        print("   3. Nhấn Đăng nhập để vào hệ thống PIM.")
        print("   4. Khi đã vào được trang https://pim.frt.vn/product/upc-list,")
        print("   5. Quay lại đây và nhấn phím [ENTER] để bot bắt đầu cào dữ liệu.")
        print("=" * 65)
        input("👉 Nhấn [ENTER] tại đây sau khi đăng nhập thành công: ")
    else:
        print("[✓] Đã sẵn sàng trên hệ thống PIM.")


def close_drawer(page):
    """
    Đóng triệt để drawer và xóa bỏ mask che màn hình.
    """
    try:
        page.keyboard.press("Escape")
        time.sleep(0.3)
    except Exception:
        pass

    try:
        close_btn = page.locator(".ant-drawer-close, button[aria-label='Close']").first
        if close_btn.is_visible(timeout=500):
            close_btn.click(force=True)
            time.sleep(0.3)
    except Exception:
        pass

    try:
        # Nếu còn mask, nhấn Escape lần nữa
        mask = page.locator(".ant-drawer-mask").first
        if mask.is_visible(timeout=500):
            page.keyboard.press("Escape")
            time.sleep(0.3)
    except Exception:
        pass


def crawl_single_product(page, code, config):
    upc_list_url = config.get("pim_url", "https://pim.frt.vn/product/upc-list")

    if "upc-list" not in page.url:
        page.goto(upc_list_url, wait_until="domcontentloaded", timeout=30000)
        time.sleep(1.5)

    # 1. Đảm bảo đóng mọi drawer và mask cũ
    close_drawer(page)

    # 2. Nhập mã vào ô tìm kiếm
    search_box = page.locator(".ant-input-search input, input#term, input[placeholder*='Tìm']").first
    search_box.click(force=True)
    search_box.fill("")
    search_box.fill(str(code))
    time.sleep(0.3)
    search_box.press("Enter")

    try:
        search_btn = page.locator(".ant-input-search-button, button:has-text('Tìm')").first
        if search_btn.is_visible(timeout=500):
            search_btn.click()
    except Exception:
        pass

    time.sleep(2.5)

    # 3. Click vào ô Tên UPC (class table__td--name hoặc td.truncate)
    upc_name_el = page.locator(".table__td--name, td.truncate").first
    if not upc_name_el.is_visible(timeout=4000):
        upc_name_el = page.locator("tbody tr td:nth-child(4)").first

    if not upc_name_el.is_visible(timeout=3000):
        return None

    upc_name_el.click()
    time.sleep(2)

    # 4. Đọc bảng thông tin bên phải (Right Drawer)
    extracted_name = None
    right_drawer = page.locator(".ant-drawer-right .ant-drawer-content, .ant-drawer-content-wrapper").first
    if right_drawer.is_visible(timeout=4000):
        drawer_txt = right_drawer.inner_text()
        extracted_name = parse_product_name_from_drawer(drawer_txt)

    if not extracted_name:
        raw = upc_name_el.inner_text().strip()
        extracted_name = raw if raw else None

    # 5. Đóng drawer ngay sau khi lấy xong
    close_drawer(page)

    return extracted_name


def run():
    print("=" * 65)
    print("      BOT TỰ ĐỘNG THU THẬP THÔNG TIN PIM.FRT.VN")
    print("=" * 65)

    config = load_config()
    excel_path = choose_excel_file(config)
    if not excel_path:
        print("[!] Không có file Excel để xử lý.")
        return

    print(f"\n[*] File Excel: {excel_path.resolve()}")
    print(f"[*] Cột mã sản phẩm: Cột {config.get('code_column', 'B')}")
    print(f"[*] Cột kết quả ghi: Cột {config.get('result_column', 'C')}")
    print(f"[*] Link xuất phát:  {config.get('pim_url', 'https://pim.frt.vn/product/upc-list')}")
    print("-" * 65)

    backup_excel(excel_path)

    try:
        wb = openpyxl.load_workbook(excel_path)
    except PermissionError:
        print(f"\n[!] CẢNH BÁO: File '{excel_path.name}' đang mở trong Excel!")
        print("    Vui lòng đóng file Excel lại rồi chạy lại bot.")
        return
    except Exception as e:
        print(f"[!] Không thể mở file Excel: {e}")
        return

    sheet_name = config.get("sheet_name", "")
    ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active
    print(f"[*] Đang làm việc trên Sheet: '{ws.title}'")

    code_col = config.get("code_column", "B").upper()
    result_col = config.get("result_column", "C").upper()
    start_row = int(config.get("start_row", 2))
    pim_url = config.get("pim_url", "https://pim.frt.vn/product/upc-list")
    delay = float(config.get("request_delay_seconds", 1.5))

    USER_DATA_DIR.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        print("\n[*] Đang khởi động trình duyệt...")
        browser_context = p.chromium.launch_persistent_context(
            user_data_dir=str(USER_DATA_DIR),
            headless=False,
            viewport={"width": 1366, "height": 768},
            args=["--start-maximized"]
        )

        page = browser_context.pages[0] if browser_context.pages else browser_context.new_page()

        print(f"[*] Đang mở: {pim_url}")
        try:
            page.goto(pim_url, wait_until="domcontentloaded", timeout=30000)
        except Exception as e:
            print(f"[!] Chú ý tải trang: {e}")

        # Kiểm tra đăng nhập
        wait_for_user_login(page)

        total_rows = ws.max_row
        print(f"\n[*] Bắt đầu xử lý từ hàng {start_row} đến hàng {total_rows}...")

        success_count = 0
        skip_count = 0
        fail_count = 0

        for row_idx in range(start_row, total_rows + 1):
            code_cell = ws[f"{code_col}{row_idx}"]
            result_cell = ws[f"{result_col}{row_idx}"]

            code = code_cell.value
            if code is None or str(code).strip() == "":
                continue

            code_str = str(code).strip()
            if code_str.endswith(".0") and code_str[:-2].isdigit():
                code_str = code_str[:-2]

            current_result = result_cell.value
            if current_result and str(current_result).strip() != "" and str(current_result).strip() != "Trang chủ":
                print(f"[=] Hàng {row_idx:3d}: Mã {code_str:<12} -> Đã có tên: '{current_result}', bỏ qua.")
                skip_count += 1
                continue

            print(f"[*] Hàng {row_idx:3d}: Mã {code_str:<12}", end="", flush=True)

            try:
                product_name = crawl_single_product(page, code_str, config)
            except Exception as e:
                print(f" [Lỗi: {e}]", end="")
                product_name = None

            if product_name:
                result_cell.value = product_name
                print(f" -> [✓] TÊN: '{product_name}'")
                success_count += 1
            else:
                print(f" -> [x] Không tìm thấy")
                fail_count += 1

            # Lưu file Excel sau mỗi dòng thành công
            try:
                wb.save(excel_path)
            except PermissionError:
                print("\n[!] File Excel đang mở bởi phần mềm khác! Vui lòng đóng lại.")
                input("    Nhấn [ENTER] sau khi đã đóng file: ")
                wb.save(excel_path)
            except Exception as e:
                print(f"\n[!] Lỗi lưu Excel: {e}")

            time.sleep(delay)

        browser_context.close()

    print("\n" + "=" * 65)
    print("🎉 HOÀN THÀNH!")
    print(f"   - Thành công:     {success_count}")
    print(f"   - Bỏ qua đã có:   {skip_count}")
    print(f"   - Không tìm thấy: {fail_count}")
    print(f"   - File kết quả:   {excel_path.resolve()}")
    print("=" * 65)


if __name__ == "__main__":
    run()
