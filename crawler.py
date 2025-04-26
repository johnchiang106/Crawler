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
from datetime import datetime

# ==== CONFIGURATIONS ====
# company_idx = [2330, 2303, 3515]  # Your stock codes
company_idx = []
output_file = 'crawler.xlsx'
years = 5  # Can be set to 1~7

company_input = input("請輸入股票代號，並以空格隔開(如2330 2303 3515): \n")
for code in company_input.split():
    try:
        company_idx.append(int(code))
    except ValueError:
        print(f"Error: '{code}' is not a valid number.")
        input("Press Enter to exit...")
        exit(1)

# ==== SETUP ====
current_time = datetime.now()
work_sheet = current_time.strftime("%H%M %Y-%m-%d")
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
row1, row2, row3 = 2, years+1, 2

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
        df.to_excel(writer, sheet_name=str(work_sheet), index=False)
except Exception as e:
    print(f"Error writing '{output_file}': {e}\nPlease close the Excel file and try again.")
    driver.quit()
    exit(0)

# ==== START CRAWLING ====
print("\nStart crawling:")
for index, company_id in enumerate(company_idx):
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
    company_names.extend([None] * years)

    # Parse basic info table
    soup = BeautifulSoup(driver.page_source, "html.parser")
    rows = soup.find("table", class_="t0").find_all("tr")
    res_basic = []
    for count, tr in enumerate(rows):
        if count in (1, 2): continue
        tds = tr.find_all('td')
        row = [td.text.strip() for col, td in enumerate(tds) if (col != 1 and col < years+2) and td.text.strip()]
        if row: res_basic.append(row)
        if count > 3: break
    df_basic = pd.DataFrame(res_basic)

    ## Step 2: Company Detail Info (PER/PBR)
    driver.get(url_detail)

    # Wait and click PER/PBR drop-down
    dropdown = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='txtFinDetailLoading']/preceding-sibling::*[1]")))
    options_in_dropdown = dropdown.find_elements(By.TAG_NAME, "option")
    options_in_dropdown[2].click()
    time.sleep(1)  # extra wait after clicking

    # Parse detail table
    soup = BeautifulSoup(driver.page_source, "html.parser")
    rows = soup.find("table", id="tblDetail").find_all("tr")
    res_detail = [["最高PER", "最低PER"]]
    for count, tr in enumerate(rows):
        if count <= 2: continue
        tds = tr.find_all('td')
        row = [td.text.strip() for col, td in enumerate(tds) if col in (10, 11) and td.text.strip()]
        if row: res_detail.append(row)
        if count > years+1: break
    df_detail = pd.DataFrame(res_detail).transpose()

    # Combine dataframes
    df_combined = pd.concat([df_basic, df_detail])
    df_combined.set_index(df_combined.columns[0], inplace=True)
    df_combined = df_combined.transpose().astype(convert_dict)
    df = pd.concat([df, df_combined])
    df.loc[len(df)+1] = pd.Series(dtype='float64')

    ## Step 3: Calculations
    yminl1 = f"=ROUND(MIN(F{row1}:F{row2}),1)"
    dminl2 = f"=ROUND(MIN(D{row1}:D{row2}),1)"
    yav1 = f"=ROUND(AVERAGE(F{row1}:F{row2}),1)"
    dav2 = f"=ROUND(AVERAGE(D{row1}:D{row2}),1)"
    max_4 = f"=ROUND(MAX(E{row1}:E{row2}),1)"

    l1_div_5 = f"=ROUND((Q{row3}-J{row3})/5+J{row3},1)"
    l2_div_5 = f"=ROUND(L{row3}*2-J{row3},1)"
    l3_div_5 = f"=ROUND(O{row3}*2-L{row3},1)"

    row1 += years+1
    row2 += years+1
    row3 += 1

    calculations.append([company_name, yminl1, dminl2, l1_div_5, yav1, dav2, l2_div_5, l3_div_5, max_4])

# ==== Save Results ====
df_cn = pd.DataFrame(company_names)
col_names = ["公司", "YMINL1", "DMINL2", "L1/5", "YAV1", "DAV2", "L2/5", "L3/5", "4最高"]
df_cal = pd.DataFrame(calculations, columns=col_names)

with pd.ExcelWriter(output_file, mode='a', if_sheet_exists='overlay') as writer:
    df.to_excel(writer, sheet_name=str(work_sheet), index=False, startrow=0, startcol=1)
    df_cn.to_excel(writer, sheet_name=str(work_sheet), header=None, index=False, startrow=1)
    df_cal.to_excel(writer, sheet_name=str(work_sheet), index=False, startrow=0, startcol=8)

# ==== Cleanup ====
driver.quit()

print("✅ Crawling finished!")
