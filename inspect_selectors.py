import json
from pathlib import Path
from playwright.sync_api import sync_playwright

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.json"
USER_DATA_DIR = BASE_DIR / "browser_profile"


def inspect():
    print("=" * 65)
    print("    CÔNG CỤ HỖ TRỢ XÁC ĐỊNH SELECTOR (GIAO DIỆN PIM.FRT.VN)")
    print("=" * 65)

    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        config = json.load(f)

    pim_url = config.get("pim_url", "https://pim.frt.vn")

    with sync_playwright() as p:
        browser_context = p.chromium.launch_persistent_context(
            user_data_dir=str(USER_DATA_DIR),
            headless=False,
            viewport={"width": 1366, "height": 768},
            args=["--start-maximized"]
        )
        page = browser_context.pages[0] if browser_context.pages else browser_context.new_page()
        page.goto(pim_url)

        print("\n1. Trình duyệt đã mở.")
        print("2. Bạn hãy đăng nhập vào PIM nếu chưa đăng nhập.")
        print("3. Thử tìm kiếm 1 mã sản phẩm bất kỳ trên PIM để kết quả hiển thị ra màn hình.")
        input("\n👉 Sau khi đã thấy kết quả trên màn hình, hãy nhấn [ENTER] tại đây: ")

        # Lấy thông tin các bảng hoặc danh sách sản phẩm hiển thị trên trang
        print("\n[*] Đang quét các thành phần văn bản trên màn hình...")

        table_rows = page.locator("table tbody tr").all()
        if table_rows:
            print(f"\n[+] Tìm thấy bảng với {len(table_rows)} dòng dữ liệu.")
            first_row = table_rows[0]
            cells = first_row.locator("td").all()
            print(f"    Chi tiết các cột ở dòng đầu tiên:")
            for idx, cell in enumerate(cells, 1):
                txt = cell.inner_text().strip().replace("\n", " ")
                print(f"    - Cột {idx} (selector: 'table tbody tr:first-child td:nth-child({idx})'): {txt[:50]}")
        else:
            print("[-] Không phát hiện thẻ <table> tiêu chuẩn. Đang quét các thẻ chứa tiêu đề/tên:")
            headings = page.locator("h1, h2, h3, h4, .title, .product-name, .name, [class*='name']").all()
            for h in headings[:10]:
                txt = h.inner_text().strip()
                if txt:
                    tag = h.evaluate("el => el.tagName.toLowerCase() + (el.className ? '.' + el.className.split(' ').join('.') : '')")
                    print(f"    - Phần tử <{tag}>: {txt[:60]}")

        print("\n" + "=" * 65)
        print("Bạn có thể copy selector chính xác ở trên vào mục 'product_name' trong file config.json")
        print("=" * 65)
        input("Nhấn [ENTER] để đóng công cụ hỗ trợ...")
        browser_context.close()


if __name__ == "__main__":
    inspect()
