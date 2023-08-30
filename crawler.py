from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import pandas as pd
import time
# from datetime import date

currTime = time.time()
# currDate = date.today().strftime("%Y%m%d")

output_file = 'crawler.xlsx'
company_idx = [3515, 2303, 2049, 1101] # 輸入公司代碼，以逗號隔開
company_names = []
calculations = []
df = pd.DataFrame()
convert_dict = {'年度': int,
                '最高PER': float,
                '最低PER': float,
                '最高本益比': float,
                '最低本益比': float,
                }

try:
    with pd.ExcelWriter(output_file, mode='a', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, sheet_name=str(currTime), index=False)
except Exception as e:
    print(f"Error writing '{output_file}': {e}\nPlease close {output_file} file")
    exit(0)

service = Service(executable_path=r'./chromedriver-win64/chromedriver.exe')
# option.add_argument("-headless")
option = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options = option)

for idx in range(len(company_idx)):
    url0 = "https://goodinfo.tw/tw/StockBzPerformance.asp?STOCK_ID=" + str(company_idx[idx])
    url1 = "https://concords.moneydj.com/z/zc/zca/zca_" + str(company_idx[idx]) + ".djhtm"

    driver.get(url1)
    title_whole = driver.find_element(By.XPATH, "//*[@class='t10']")
    title = title_whole.text.split(' ')
    company_name = title[0] + " " + title[1]
    print(company_name)
    company_names.append(company_name)
    company_names.extend([None, None, None, None, None, None])

    root = BeautifulSoup(driver.page_source, "html.parser")
    rows = root.find( "table", {"class": "t0"}).find_all("tr")
    res = []
    for count, tr in enumerate(rows):
        if count == 1 or count == 2:
            continue
        td = tr.find_all('td')
        row = [tr.text.strip() for col, tr in enumerate(td) if (col != 1 and col != 8) and tr.text.strip()]
        if row:
            res.append(row)
        if count > 3:
            break

    df2 = pd.DataFrame(res)

    driver.get(url0)
    PERPBROption = driver.find_element(By.XPATH, "//*[@id='txtFinDetailLoading']/preceding-sibling::*[1]/select/option[3]")
    # print(PERPBROption.text)
    PERPBROption.click()
    time.sleep(3)

    root = BeautifulSoup(driver.page_source, "html.parser")
    rows = root.find( "table", {"id": "tblDetail"}).find_all("tr")
    res = [["最高PER", "最低PER"]]
    for count, tr in enumerate(rows):
        if count <= 2:
            continue
        td = tr.find_all('td')
        row = [tr.text.strip() for col, tr in enumerate(td) if (col == 10 or col == 11) and tr.text.strip()]
        if row:
            res.append(row)
        if count > 7:
            break
        
    df2 = pd.concat([df2, pd.DataFrame(res).transpose()])
    df2.set_index(df2.columns[0], inplace=True)
    df2 = df2.transpose().astype(convert_dict)

    df = pd.concat([df, df2])
    df.loc[len(df)+1] = pd.Series(dtype='float64')

    yminl1 = df2["最低PER"].min()
    dminl2 = df2["最低本益比"].min()
    yav1 = df2["最低PER"].mean()
    dav2 = df2["最低本益比"].mean()
    max_4 = df2["最高PER"].max()
    l1_div_5 = (max_4 - yminl1) / 5 + yminl1
    l2_div_5 = l1_div_5 * 2 - yminl1
    l3_div_5 = l2_div_5 * 2 - l1_div_5

    yminl1 = round(yminl1, 1)
    dminl2 = round(dminl2, 1)
    yav1 = round(yav1, 1)
    dav2 = round(dav2, 1)
    max_4 = round(max_4, 1)
    l1_div_5 = round(l1_div_5, 1)
    l2_div_5 = round(l2_div_5, 1)
    l3_div_5 = round(l3_div_5, 1)

    calculations.append([company_name, yminl1, dminl2, l1_div_5, yav1, dav2, l2_div_5, l3_div_5, max_4])

df_cn = pd.DataFrame(company_names)
col_names = ["公司", "YMINL1", "DMINL2", "L1/5", "YAV1", "DAV2", "L2/5", "L3/5", "4最高"]
df_cal = pd.DataFrame(calculations, columns=col_names)

with pd.ExcelWriter(output_file, mode='a', if_sheet_exists='overlay') as writer:
    df.to_excel(writer, sheet_name=str(currTime), index=False, startrow=0, startcol=1)
    df_cn.to_excel(writer, sheet_name=str(currTime), header=None, index=False, startrow=1)
    df_cal.to_excel(writer, sheet_name=str(currTime), index=False, startrow=0, startcol=8)