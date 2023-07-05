import time, openpyxl
from datetime import datetime
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC

def makeWebDriver():
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--start-maximized")       # 視窗最大化
    chrome_options.add_argument('--headless')        # 背景执行
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    browser = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

    return browser

def select_category():
    try:
        _browser.switch_to.frame(_browser.find_element(By.XPATH, '//*[@id="iMARK"]'))
    except:
        pass
    WebDriverWait(_browser, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@class="main_form"]')))
    tbody = _browser.find_element(By.XPATH, '//*[@class="main_form"]')
    checkbox = tbody.find_elements(By.XPATH, '//*[@type="checkbox"]')
    checkbox = checkbox[11]
    _browser.execute_script("arguments[0].click();", checkbox)
    WebDriverWait(_browser, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@data-role="listview"]')))

# 找出符合權證(排除名字含有"售"、"反")
def count_warrant():
    select_category()
    status = 1
    count = 0
    print("-----當前符合權證-----")
    while status == 1:
        table = _browser.find_element(By.XPATH, '//*[@data-role="listview"]')
        warrant_rows = table.find_elements(By.TAG_NAME,'tr')
        for row in range(len(warrant_rows)):
            find_warrant = warrant_rows[row]
            try:
                _browser.switch_to.frame(_browser.find_element(By.XPATH, '//*[@id="iMARK"]'))
            except:
                pass
            WebDriverWait(_browser, 15).until(EC.element_to_be_clickable((By.XPATH, '//*[contains(@href, "/EDWebSite/Controllers/WarrantRoute.aspx?")]')))
            warrant_attr = find_warrant.find_elements(By.XPATH, '//*[contains(@href, "/EDWebSite/Controllers/WarrantRoute.aspx?")]//*[contains(text(), "購") and not(contains(text(), "反"))]')
        for attr in range(len(warrant_attr)):
            count = count+1
            warrant_name = warrant_attr[attr]
            _browser.execute_script("arguments[0].scrollIntoView();", warrant_name)
            print(str(count)+"、"+warrant_name.text)

        if row == 19:
            status = 0

    return count

# 逐一分析符合權證
def find_warrant(count):
    status = 1
    times = 1
    print("-----分析結果-----")
    while status == 1:
        table = _browser.find_element(By.XPATH, '//*[@data-role="listview"]')
        warrant_rows = table.find_elements(By.TAG_NAME,'tr')
        for row in range(len(warrant_rows)):
            find_warrant = warrant_rows[row]
            WebDriverWait(_browser, 15).until(EC.element_to_be_clickable((By.XPATH, '//*[contains(@href, "/EDWebSite/Controllers/WarrantRoute.aspx?")]')))
            warrant_info = find_warrant.find_elements(By.XPATH, '//*[contains(@href, "/EDWebSite/Controllers/WarrantRoute.aspx?")]')
            warrant_attr = find_warrant.find_elements(By.XPATH, '//*[contains(@href, "/EDWebSite/Controllers/WarrantRoute.aspx?")]//*[contains(text(), "購") and not(contains(text(), "反"))]')
            for info in range(len(warrant_info)):
                for attr in range(len(warrant_attr)):
                    try:
                        _browser.switch_to.frame(_browser.find_element(By.XPATH, '//*[@id="iMARK"]'))
                    except:
                        pass
                    find_info = warrant_info[info]
                    warrant_name = warrant_attr[attr]
                    if find_info.text == warrant_name.text:
                        _browser.execute_script("arguments[0].scrollIntoView();", warrant_name)
                        warrant_url = warrant_info[info-1]
                        url = warrant_url.get_attribute("href")
                        _browser.execute_script("window.open('')")
                        _browser.switch_to.window(_browser.window_handles[1])
                        _browser.get(url)
                        try:
                            WebDriverWait(_browser, 10).until(EC.presence_of_element_located((By.XPATH, '//*[text()="權證基本資料"]')))
                            basic_info = _browser.find_element(By.XPATH, '//label[text()="權證基本資料"]')
                            _browser.execute_script("arguments[0].click();", basic_info)
                        except:
                            pass
                        WebDriverWait(_browser, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ifWarrantAnalyzer"]')))
                        _browser.switch_to.frame(_browser.find_element(By.XPATH, '//*[@id="ifWarrantAnalyzer"]'))
                        time.sleep(5)
                        times = times+1
                        analysis_data(count, times)
                        _browser.close()
                        _browser.switch_to.window(_browser.window_handles[0])

            if times >= attr:
                status = 0
                break

# 分析資料
def analysis_data(count, times):
    data = []
    WebDriverWait(_browser, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="underlyingData"]')))
    WebDriverWait(_browser, 15).until(EC.presence_of_element_located((By.XPATH, '//*[@id="warrantData"]')))
    warrant_detail = _browser.find_element(By.XPATH, '//*[@id="warrantDataDetail"]')
    detail_tr = warrant_detail.find_elements(By.TAG_NAME, 'tr')
    flux_tr = detail_tr[3]
    flux_td = flux_tr.find_elements(By.TAG_NAME, 'td')
    warrant_flux = flux_td[7]
    rate_tr = detail_tr[4]
    rate_td = rate_tr.find_elements(By.TAG_NAME, 'td')
    warrant_rate = rate_td[7]
    # 取前一日在外流通張數低於500或是在外流通率低於10％
    if int((warrant_flux.text).replace(",", "")) < 500 or int(float((warrant_rate.text).replace("%", ""))) < 20:
        warrant_data = _browser.find_element(By.XPATH, '//*[@id="warrantData"]')
        data_td = warrant_data.find_elements(By.TAG_NAME, 'td')
        warrant_code = data_td[0]
        data.append("權證代碼："+warrant_code.text)
        warrant_name = data_td[1]
        data.append("權證名稱："+warrant_name.text)
        warrant_price = data_td[6]
        data.append("當前價格："+warrant_price.text)

        print(data)

# 寫資料 目前是用print的方式顯示，故寫入excel暫不用
# def write_data(count, times, data):
#     row = times+2
#     for column in range(1, len(data)+1):
#         w_sheet.cell(row, column).value = data[column-1]
#     write_times = write_times+1

#     if write_times == count:
#         w_book.save('analysis_warrant.xlsx')
#         print("資料寫入完成")

if __name__ == "__main__":
    while True:
        if int(time.strftime("%H%M", time.localtime())) == 1330:
            print("-----當前分析時間：1330，結束分析-----")
            break
        elif int(time.strftime("%M", time.localtime()))%10 == 0:
            url = "https://warrant.kgi.com/EDWebSite/Views/StrategyCandidate/MarketStatisticsIframe.aspx"
            write_times = 0
            w_book = openpyxl.Workbook()
            w_sheet = w_book.worksheets[0]
            w_sheet.merge_cells('A1:G1')
            warrant_info = ["權證代碼", "權證名稱", "當前價格", "在外流通量(%)", "凱基網址"]
            w_sheet['A1'] = "當前分析時間：" + str(datetime.now())
            for column in range(1, len(warrant_info)+1):
                w_sheet.cell(2, column).value = warrant_info[column-1]

            _browser = makeWebDriver()
            print("當前時間"+time.strftime("%Y-%m-%d %H:%M:%S" , time.localtime()))
            _browser.get(url)
            count = count_warrant()
            find_warrant(count)
            _browser.quit()
            print("-----本次分析結束-----"+"\n")