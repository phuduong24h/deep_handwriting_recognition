from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time
import pandas as pd

# Cấu hình Selenium headless
options = Options()
options.add_argument("--headless")
options.add_argument("--disable-gpu")
options.add_argument("--window-size=1920,1080")

driver = webdriver.Chrome(options=options)

all_data = []
so_trang_muon_crawl = 5  # ← Bạn có thể tăng lên bao nhiêu trang tùy ý

for page in range(1, so_trang_muon_crawl + 1):
    url = f"https://drugbank.vn/danh-sach-thuoc?page={page}"
    print(f"Đang xử lý trang {page}...")
    driver.get(url)
    time.sleep(2)  # Đợi trang tải

    rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")

    for row in rows:
        tds = row.find_elements(By.TAG_NAME, "td")
        if len(tds) >= 3:
            ten_thuoc = tds[0].text.strip()
            dang_bao_che = tds[1].text.strip()
            nha_san_xuat = tds[2].text.strip()
            all_data.append({
                "Số đăng kí": ten_thuoc,
                "Tên thuốc": dang_bao_che,
                "Hoạt chất": nha_san_xuat
            })

driver.quit()

# Ghi dữ liệu ra file Excel
df = pd.DataFrame(all_data)
df.to_excel("danh_sach_thuoc_nhieu_trang.xlsx", index=False)
print("✅ Đã crawl xong và lưu file Excel!")
