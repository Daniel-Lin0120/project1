import re
import time
import random
import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from fake_useragent import UserAgent

# 🛠 產生隨機 User-Agent
ua = UserAgent()
random_user_agent = ua.random

# 🛠 啟動 Selenium WebDriver
options = webdriver.ChromeOptions()
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("--incognito")
options.add_argument(f"user-agent={random_user_agent}")

driver = webdriver.Chrome(service=webdriver.ChromeService(ChromeDriverManager().install()), options=options)
wait = WebDriverWait(driver, 10)

# 🛠 讀取 Excel 檔案
input_file = "公司列表.xlsx"  
output_file = "公司列表_完整資訊.xlsx"
df = pd.read_excel(input_file)

# 🛠 讀取郵遞區號對照表
zipcode_file = "郵遞區號.xlsx"
zipcode_df = pd.read_excel(zipcode_file)
zipcode_mapping = dict(zip(zipcode_df["區域"], zipcode_df["郵遞區號"].astype(str)))

# 🛠 確保有 "公司名稱" 欄位
if "公司名稱" not in df.columns:
    print("❌ Excel 檔案缺少 '公司名稱' 欄位！")
    exit()

# 🛠 查詢統一編號的函式
def get_company_id(company_name):
    search_url = f"https://tw.piliapp.com/vat-calculator/tw/search/?q={company_name}"
    driver.get(search_url)
    time.sleep(random.uniform(2, 4))  # 減少等待時間，但仍保留隨機性

    try:
        company_id = driver.find_element(By.XPATH, "//table/tbody/tr[1]/td[1]").text.strip()
    except:
        company_id = "(查無資料)"

    return company_id

# 🛠 查詢公司詳細資訊的函式
def get_company_details(company_id):
    if company_id == "(查無資料)":
        return {"公司全名": "(查無資料)", "地址": "(查無資料)", "總經理": "(查無資料)", "董事長": "(查無資料)", "電話": "(查無資料)", "信箱": "(查無資料)", "第一位經理人": "(查無資料)"}

    company_url = f"https://www.twincn.com/item.aspx?no={company_id}"
    driver.get(company_url)
    time.sleep(random.uniform(1, 2))  # 避免過快存取

    def extract_data(xpath, default_value="(查無資料)"):
        try:
            return driver.find_element(By.XPATH, xpath).text.strip().split("\n")[0]  # 取第一行
        except:
            return default_value

    # 🛠 取得第一位經理人
    def get_first_manager():
        try:
            first_manager = driver.find_element(By.XPATH, "(//table)[4]/tbody/tr[last()]/td[1]").text.strip()
            # **檢查是否為人名 (排除帶括號、商標、公司名稱等內容)**
            if not re.match(r"^[\u4e00-\u9fa5·]{2,4}$", first_manager):  # 只允許中文姓名
                return "(查無資料)"
            return first_manager
        except:
            return "(查無資料)"

    address = extract_data("//td[strong[contains(text(),'公司所在地')]]/following-sibling::td")
    zipcode_prefix = ""
    for area, zipcode in zipcode_mapping.items():
        if area in address:
            zipcode_prefix = zipcode[:3]  # 取前三碼
            break

    return {
        "公司全名": extract_data("//td[strong[contains(text(),'公司名稱')]]/following-sibling::td"),
        "地址": f"{zipcode_prefix} {address}" if zipcode_prefix else address,
        "總經理": get_first_manager(),  # 抓取經理人姓名
        "董事長": extract_data("//td[contains(text(),'董事長')]/following-sibling::td"),
        "電話": extract_data("//td[strong[contains(text(),'電話')]]/following-sibling::td"),
        "信箱": extract_data("//td[strong[contains(text(),'Mail')]]/following-sibling::td"),
    }

# 🛠 開始查詢
df["統一編號"] = df["公司名稱"].apply(get_company_id)
df["公司全名"] = ""
df["地址"] = ""
df["總經理"] = ""
df["董事長"] = ""
df["電話"] = ""
df["信箱"] = ""

for index, row in df.iterrows():
    if row["統一編號"] != "(查無資料)":
        details = get_company_details(row["統一編號"])
        df.at[index, "公司全名"] = details["公司全名"]
        df.at[index, "地址"] = details["地址"]
        df.at[index, "總經理"] = details["總經理"]
        df.at[index, "董事長"] = details["董事長"]
        df.at[index, "電話"] = details["電話"]
        df.at[index, "信箱"] = details["信箱"]

# 🛠 關閉瀏覽器
driver.quit()

# 🛠 儲存結果到 Excel
df.to_excel(output_file, index=False, engine="openpyxl")
print(f"✅ 查詢完成，結果已存入 {output_file}")
