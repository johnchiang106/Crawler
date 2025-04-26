from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import pandas as pd
import time

# ==== CONFIGURATIONS ====
company_idx = [2330]  # Your stock codes
output_file = 'crawler.xlsx'

# ==== SETUP ====
currTime = time.time()
df = pd.DataFrame()
company_names = []
calculations = []
convert_dict = {
    '年度': int,
    '最高PER': float,
    '最低PER': float,
    '最高本益比': float,
    '最低本益比': float,
}

# ==== Chrome Options ====
options = Options()
# options.add_argument("--headless")  # Uncomment if you don't want the browser window
options.add_argument("--ignore-certificate-errors")
options.add_argument("--log-level=3")
options.add_experimental_option('excludeSwitches', ['enable-logging'])

# ==== Launch Browser ====
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wait = WebDriverWait(driver, 10)

# ==== Ensure output file is ready ====
try:
    with pd.ExcelWriter(output_file, mode='a', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, sheet_name=str(currTime), index=False)
except Exception as e:
    print(f"Error writing '{output_file}': {e}\nPlease close the Excel file and try again.")
    driver.quit()
    exit(0)

# ==== START CRAWLING ====
print("Start crawling:")
for company_id in company_idx:
    # URLs
    url_basic = f"https://concords.moneydj.com/z/zc/zca/zca_{company_id}.djhtm"
    url_detail = f"https://goodinfo.tw/tw/StockBzPerformance.asp?STOCK_ID={company_id}"

    ## Step 1: Company Basic Info
    driver.get(url_basic)
    title_element = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "t10")))
    title_text = title_element.text.split(' ')
    company_name = f"{title_text[0]} {title_text[1]}"
    print(company_name)
    company_names.append(company_name)
    company_names.extend([None] * 6)

    # Parse basic info table
    soup = BeautifulSoup(driver.page_source, "html.parser")
    rows = soup.find("table", class_="t0").find_all("tr")
    res_basic = []
    for count, tr in enumerate(rows):
        if count in (1, 2): continue
        tds = tr.find_all('td')
        row = [td.text.strip() for col, td in enumerate(tds) if (col != 1 and col != 8) and td.text.strip()]
        if row: res_basic.append(row)
        if count > 3: break
    df_basic = pd.DataFrame(res_basic)

    ## Step 2: Company Detail Info (PER/PBR)
    driver.get(url_detail)

    # Wait and click PER/PBR drop-down
    dropdown = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='txtFinDetailLoading']/preceding-sibling::*[1]")))
    options_in_dropdown = dropdown.find_elements(By.TAG_NAME, "option")
    options_in_dropdown[2].click()
    time.sleep(3)  # extra wait after clicking

    # Parse detail table
    soup = BeautifulSoup(driver.page_source, "html.parser")
    rows = soup.find("table", id="tblDetail").find_all("tr")
    res_detail = [["最高PER", "最低PER"]]
    for count, tr in enumerate(rows):
        if count <= 2: continue
        tds = tr.find_all('td')
        row = [td.text.strip() for col, td in enumerate(tds) if col in (10, 11) and td.text.strip()]
        if row: res_detail.append(row)
        if count > 7: break
    df_detail = pd.DataFrame(res_detail).transpose()

    # Combine dataframes
    df_combined = pd.concat([df_basic, df_detail])
    df_combined.set_index(df_combined.columns[0], inplace=True)
    df_combined = df_combined.transpose().astype(convert_dict)
    df = pd.concat([df, df_combined])
    df.loc[len(df)+1] = pd.Series(dtype='float64')

    ## Step 3: Calculations
    yminl1 = df_combined["最低PER"].min()
    dminl2 = df_combined["最低本益比"].min()
    yav1 = df_combined["最低PER"].mean()
    dav2 = df_combined["最低本益比"].mean()
    max_4 = df_combined["最高PER"].max()

    l1_div_5 = (max_4 - yminl1) / 5 + yminl1
    l2_div_5 = l1_div_5 * 2 - yminl1
    l3_div_5 = l2_div_5 * 2 - l1_div_5

    # Round
    yminl1 = round(yminl1, 1)
    dminl2 = round(dminl2, 1)
    yav1 = round(yav1, 1)
    dav2 = round(dav2, 1)
    max_4 = round(max_4, 1)
    l1_div_5 = round(l1_div_5, 1)
    l2_div_5 = round(l2_div_5, 1)
    l3_div_5 = round(l3_div_5, 1)

    calculations.append([company_name, yminl1, dminl2, l1_div_5, yav1, dav2, l2_div_5, l3_div_5, max_4])

# ==== Save Results ====
df_cn = pd.DataFrame(company_names)
col_names = ["公司", "YMINL1", "DMINL2", "L1/5", "YAV1", "DAV2", "L2/5", "L3/5", "4最高"]
df_cal = pd.DataFrame(calculations, columns=col_names)

with pd.ExcelWriter(output_file, mode='a', if_sheet_exists='overlay') as writer:
    df.to_excel(writer, sheet_name=str(currTime), index=False, startrow=0, startcol=1)
    df_cn.to_excel(writer, sheet_name=str(currTime), header=None, index=False, startrow=1)
    df_cal.to_excel(writer, sheet_name=str(currTime), index=False, startrow=0, startcol=8)

# ==== Cleanup ====
driver.quit()

print("✅ Crawling finished!")
