from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import pandas as pd
import time
from datetime import date

currTime = time.time()
currDate = date.today().strftime("%Y%m%d")

service = Service(executable_path=r'./chromedriver-win64/chromedriver.exe')
# option.add_argument("-headless")
option = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options = option)

company_idx = [3515]#, 1101, 8069]

for idx in range(len(company_idx)):
    url0 = "https://goodinfo.tw/tw/StockBzPerformance.asp?STOCK_ID=" + str(company_idx[idx])
    url1 = "https://concords.moneydj.com/z/zc/zca/zca_" + str(company_idx[idx]) + ".djhtm"

    driver.get(url1)
    title_whole = driver.find_element(By.XPATH, "//*[@class='t10']")
    title = title_whole.text.split(' ')
    company_name = title[0] + " " + title[1]
    print(company_name)
    df = pd.DataFrame([[company_name]])
    with pd.ExcelWriter('crawler.xlsx', mode='a', if_sheet_exists='overlay') as writer:  
        df.to_excel(writer, sheet_name=str(currTime), header=None, index=False, startrow=idx*6, startcol=7)

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

    df = pd.DataFrame(res)
    with pd.ExcelWriter('crawler.xlsx', mode='a', if_sheet_exists='overlay') as writer:  
        df.to_excel(writer, sheet_name=str(currTime), header=None, index=False, startrow=idx*6)

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
        
    df = pd.DataFrame(res)
    df = df.transpose()

    with pd.ExcelWriter('crawler.xlsx', mode='a', if_sheet_exists='overlay') as writer:  
        df.to_excel(writer, sheet_name=str(currTime), header=None, index=False, startrow=idx*6+3)

# print(len(rows))

