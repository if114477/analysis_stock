import time, openpyxl, requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC

def makeWebDriver():
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--start-maximized")       # 視窗最大化
    # chrome_options.add_argument('--headless')        # 背景执行
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    browser = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

    return browser

def select_category():
    try:
        _browser.switch_to.frame(_browser.find_element(By.XPATH, '//*[@id="iMARK"]'))
    except:
        pass
    WebDriverWait(_browser, 50).until(EC.presence_of_element_located((By.XPATH, '//*[@class="main_form"]')))
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
    print("-----當前符合權證-----", flush=True)
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
            print(str(count)+"、"+warrant_name.text, flush=True)

        if row == 19:
            status = 0

    return count

# 逐一分析符合權證
def find_warrant(count, b_xml):
    status = 1
    times = 1
    buy_toline = []
    buy_toprt = []
    print("-----分析結果-----", flush=True)
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
                        try:
                            WebDriverWait(_browser, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ifWarrantAnalyzer"]')))
                            _browser.switch_to.frame(_browser.find_element(By.XPATH, '//*[@id="ifWarrantAnalyzer"]'))
                            time.sleep(3)
                            buy_forprt, buy_forline, b_forxml= analysis_data(buy_toprt, buy_toline, b_xml)
                        except:
                            pass
                        _browser.close()
                        _browser.switch_to.window(_browser.window_handles[0])
                        times = times+1

            # 顯示當次結果
            if times >= attr:
                buy_prt = '\n'.join(buy_forprt)
                if len(buy_prt)!=0:
                    print("大戶買進：\n"+buy_prt)
                else:
                    print("大戶買進：無")
                buy_line = '\n'.join(buy_forline)
                # sell_toStr = '\n'.join(sell)
                message = "大戶買進：\n"+buy_line
                if len(buy_line)!=0:
                    line_notify(message)
                # print("大戶賣出：\n"+sell_toStr)
                # notification(buy_toStr, sell_toStr)
                status = 0
                break
    return b_forxml

# 分析資料
def analysis_data(buy_toprt, buy_toline, b_toxml):
    buy_list = []
    WebDriverWait(_browser, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="underlyingData"]')))
    warrant_detail = _browser.find_element(By.XPATH, '//*[@id="warrantDataDetail"]')
    _browser.execute_script("arguments[0].scrollIntoView();", warrant_detail)

    # 基本資料
    detail_tr = warrant_detail.find_elements(By.TAG_NAME, 'tr')
    flux_tr = detail_tr[3]
    flux_td = flux_tr.find_elements(By.TAG_NAME, 'td')
    warrant_flux = flux_td[7].text.replace(",", "")   # 在外流通數
    rate_tr = detail_tr[4]
    rate_td = rate_tr.find_elements(By.TAG_NAME, 'td')
    warrant_total = rate_td[5].text.replace(",", "")  # 總發行張數
    warrant_rate = rate_td[7].text.replace("%", "")   # 在外流通率

    # 當前資料
    warrant_data = _browser.find_element(By.XPATH, '//*[@id="warrantData"]')
    data_td = warrant_data.find_elements(By.TAG_NAME, 'td')
    warrant_code = data_td[0].text   # 權證代碼 
    warrant_name = data_td[1].text   # 權證名稱
    warrant_price = data_td[6].text  # 權證當前價格
    warrant_vol = data_td[7].text.replace(",", "")    # 權證當前交易量

    # 取前一日在外流通張數低於1000張或是在外流通率低於10％，當作大戶買進依據
    if int(warrant_flux) < 1000 or int(float(warrant_rate)) < 10:
        buy_toprt.append(warrant_code+" "+warrant_name+" 當前價格："+warrant_price+" 交易量："+warrant_vol+" 總發行："+warrant_total+" 在外流通："+warrant_flux)
        buy_list.extend([warrant_code, warrant_name, warrant_price, warrant_vol, warrant_total, warrant_flux])
        
        condition = 1
        # 資料存在時，本次的交易量大於1000且大於前一次的交易量100時才發出通知
        for b in range(0, len(b_toxml), 6):
            if b_toxml[b] == buy_list[0]:
                condition = 0
                if int(buy_list[3])-int(b_toxml[b+3]) >= 100:
                    buy_toline.append(warrant_code+" "+warrant_name+" 當前價格："+warrant_price+" 交易量："+warrant_vol+" 總發行："+warrant_total+" 在外流通："+warrant_flux)
                    break
        # 資料不存在時，本次的交易量大於1000才發出通知
        if int(buy_list[3]) < 1000:
            condition = 0

        if condition == 1:
            buy_toline.append(warrant_code+" "+warrant_name+" 當前價格："+warrant_price+" 交易量："+warrant_vol+" 總發行："+warrant_total+" 在外流通："+warrant_flux)

        add = 1
        # 資料存在時，覆蓋上新的資料
        for b in range(0, len(b_toxml), 6):
            if b_toxml[b] == buy_list[0]:
                add = 0
                for index in range(0, 6):
                    b_toxml[b+index] = buy_list[index]
        # 資料不存在時，本次的交易量大於1000才存入陣列
        if int(buy_list[3]) < 1000:
            add = 0

        if add == 1:
            b_toxml.extend([warrant_code, warrant_name, warrant_price, warrant_vol, warrant_total, warrant_flux])

    # 取前一日在外流通張數高於10000張或是在外流通率高80％，當作大戶賣出依據
    # if int((warrant_flux.text).replace(",", "")) > 10000 or int(float((warrant_rate.text).replace("%", ""))) > 80:
    #     sell.append(warrant_code.text+" "+warrant_name.text+" 當前價格："+warrant_price.text+" 交易量："+warrant_vol.text+" 總發行："+warrant_total.text+" 在外流通："+warrant_flux.text)
    #     sell_list.extend([warrant_code.text, warrant_name.text, warrant_price.text, warrant_vol.text, warrant_total.text, warrant_flux.text])
    #     # 判斷符合權證是否存在於陣列中，如果沒有則存進陣列 用途是寫進excel
    #     if len(s_toxml) == 0:
    #         s_toxml.extend([warrant_code.text, warrant_name.text, warrant_price.text, warrant_vol.text, warrant_total.text, warrant_flux.text])
    #     else:
    #         for s in range(0, len(s_toxml), 6):
    #             if s_toxml[s] == sell_list[0]:
    #                 for index in range(0, 5):
    #                     s_toxml[s] = sell_list[index]
    #             else:
    #                 add = 1
    #         if add == 1:        
    #             s_toxml.extend([warrant_code.text, warrant_name.text, warrant_price.text, warrant_vol.text, warrant_total.text, warrant_flux.text])

    return buy_toprt, buy_toline, b_toxml

# 發送line通知
def line_notify(msg):
    Line_Notify_Account = {'token':'bKJklUOQUc5A0FM8fxzY9OgxNO0XRUO3TURcOV4sDCK'} # eVjVO4y8jiQTCwHkGtzuOyMLZqqiZKUklr20dg8bcWJ

    headers = {"Authorization": "Bearer " + Line_Notify_Account['token'],
               "Content-Type" : "application/x-www-form-urlencoded"}

    params = {"message":msg}

    r = requests.post("https://notify-api.line.me/api/notify", headers=headers, params=params)

# 寫資料 儲存每日結果至excel
def write_data(buy):
    # 建立日結表單
    w_book = openpyxl.Workbook()
    ws = w_book["Sheet"]
    w_book.remove(ws)
    w_book.create_sheet("買進", 0)
    # w_book.create_sheet("賣出", 1)

    for sheet in w_book:
        sheet.column_dimensions['A'].width=12
        sheet.column_dimensions['B'].width=25
        sheet.column_dimensions['C'].width=10
        sheet.column_dimensions['D'].width=15
        sheet.column_dimensions['E'].width=15
        sheet.column_dimensions['F'].width=20

    # 處理excel
    buy_sheet = w_book["買進"]
    warrant_info = ["權證代碼", "權證名稱", "價格", "交易量", "總發行數量", "在外流通張數"]
    # 寫入標籤
    for title in range(1, len(warrant_info)+1):
        buy_sheet.cell(1, title).value = warrant_info[title-1]
    # 寫入資料
    for column in range(2, (len(buy)//6)+2):
        index = 12
        for row in range(1, 7):
            buy_sheet.cell(column, row).value = buy[(column*6)-index]
            index = index-1

    # sell_sheet = w_book["賣出"]
    # # 寫入標籤
    # for title in range(1, len(warrant_info)+1):
    #     sell_sheet.cell(1, title).value = warrant_info[title-1]
    # # 寫入資料    
    # for column in range(2, (len(sell)//6)+2):
    #     index = 12
    #     for row in range(1, 7):
    #         sell_sheet.cell(column, row).value = sell[(column*6)-index]
    #         index = index-1

    w_book.save(str(time.strftime("%Y%m%d", time.localtime()))+'_日結.xlsx')

if __name__ == "__main__":
    _browser = makeWebDriver()
    buy_xml=[]
    url = "https://warrant.kgi.com/EDWebSite/Views/StrategyCandidate/MarketStatisticsIframe.aspx"
    while True:
        now = int(time.strftime("%H%M", time.localtime()))
        if (now>=901 or now<=1320):
            print("當前時間"+time.strftime("%Y-%m-%d %H:%M:%S" , time.localtime()), flush=True)
            _browser.get(url)
            count = count_warrant()
            buy_daily = find_warrant(count, buy_xml)
            _browser.execute_script("window.open('')")
            _browser.switch_to.window(_browser.window_handles[0])
            _browser.close()
            _browser.switch_to.window(_browser.window_handles[0])
            print("-----本次分析結束-----\n", flush=True)
        elif now >= 1330:
            _browser.quit()
            print("-----當前分析時間：1330，結束分析-----", flush=True)
            write_data(buy_daily)
            print("資料寫入完成", flush=True)
            break
        else:
            time.sleep(60)