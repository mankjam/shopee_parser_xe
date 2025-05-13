#-*- coding:utf-8 -*-
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import time
import pyodbc
from datetime import datetime
from dateutil import relativedelta
from tkinter import *
from tkinter import messagebox
from PIL import ImageTk, Image
import configparser
import uuid
import traceback
# import undetected_chromedriver as uc

# 版號
ver_num = '4.2'

config = configparser.ConfigParser()
config.read('config.ini')
# 資料庫設定
odbc_driver = ''
driver = ''
drivers = [item for item in pyodbc.drivers()]
for d in drivers:
    if 'Native Client' in d:
        driver = d
    if d.endswith(' for SQL Server'):
        odbc_driver = d
if odbc_driver:
    driver = odbc_driver
print(driver)
sql_server = config['database']['server']
sql_db = config['database']['db']
sql_uid = config['database']['uid']
sql_pwd = config['database']['pwd']

user_id = config['webpage']['user_id']
user_pwd = config['webpage']['user_pwd']

connect_string = "Driver={%s};Server=%s;Database=%s;UID=%s;PWD=%s;" % (driver, sql_server, sql_db, sql_uid, sql_pwd)

conn = pyodbc.connect(connect_string)

cursor = conn.cursor()

# 客戶編號
# 金固 097
# 正玉佳 078
# 合日昇 646
get_customer_code_sql = "SELECT detail FROM Sys_Parameter WHERE Ord = '107'"
cursor.execute(get_customer_code_sql)
result = cursor.fetchone()
product_no = result[0]
subtotal_places = 0
sql = "SELECT detail FROM Sys_Parameter WHERE Ord = '203'"
cursor.execute(sql,)
result = cursor.fetchone()
if result is not None:
    subtotal_places = int(result[0])
now = datetime.now()
build_date = now.strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
order_date = now.strftime("%Y-%m-%d ") + "00:00:00.000"
memo_string = "Sunrise(" + now.strftime("%Y/%m/%d %H:%M:%S") + ")"
nextmonth = now + relativedelta.relativedelta(months=1)
account_year = nextmonth.strftime('%Y%m')
# 出貨類別
class_no_dict = {
    '全部': '',
    '7-11': '01',
    '7-ELEVEN': '01',
    '萊爾富': '02',
    '全家': '03',
    'OK Mart': '04',
    '蝦皮宅配': '05',
    '黑貓宅急便': '06',
    '萊爾富-經濟包': '07',
    '其他寄送方式': '08',
    '蝦皮店到店': '09',
    'OK Mart, 蝦皮店到店': '09', # OK 蝦皮店到店可互寄 開啟選項後會將兩個合併成同一TAG 再判斷是不是OK 不是的話就蝦皮店到店
    '蝦皮店到店, OK Mart': '09', # OK 蝦皮店到店可互寄 開啟選項後會將兩個合併成同一TAG 再判斷是不是OK 不是的話就蝦皮店到店
    '店到家宅配': '10',
    '宅配通': '11',
    '嘉里快遞': '12',
    '全家冷凍超取': '03',
    'Other Logistics': '08',
    '蝦皮店到店 - 隔日到貨': '09',
}

class_no_dict_646 = {
    '全部': '',
    '7-11': 'A',
    '7-ELEVEN': 'A',
    '萊爾富': 'B',
    '全家': 'C',
    'OK Mart': 'G',
    '蝦皮宅配': 'D',
    '黑貓宅急便': 'E',
    '萊爾富-經濟包': 'F',
    '其他寄送方式': 'K',
    'OK Mart, 蝦皮店到店': 'G', # OK 蝦皮店到店可互寄 開啟選項後會將兩個合併成同一TAG 再判斷是不是OK 不是的話就蝦皮店到店
    '蝦皮店到店, OK Mart': 'G', # OK 蝦皮店到店可互寄 開啟選項後會將兩個合併成同一TAG 再判斷是不是OK 不是的話就蝦皮店到店
    '店到家宅配': 'H',
    '宅配通': 'I',
    '嘉里快遞': 'J',
    '全家冷凍超取': 'C',
    'Other Logistics': 'K',
    '蝦皮店到店 - 隔日到貨': 'G',
}

def CalItem():
    options = Options()
    options.add_argument('--disable-notifications')  # 取消所有的alert彈出視窗

    # options = webdriver.ChromeOptions()
    # options.add_argument('user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.198 Safari/537.36')
    options.add_argument("--disable-blink-features")
    options.add_argument("--disable-blink-features=AutomationControlled")
    # options.add_argument("--profile-directory=Default")

    # # 自動登入
    browser = webdriver.Chrome(service=Service('C:\\chromedriver-win32\\chromedriver.exe'), options=options)
    browser.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": """
            Object.defineProperty(navigator, 'webdriver', {
            get: () => undefined
            })
        """
    })
    # options = uc.ChromeOptions()
    # browser = uc.Chrome(use_subprocess=False, options=options,)
    # 利用stealth.min.js隐藏浏览器指纹特征
    # stealth.min.js下载地址：https://github.com/berstend/puppeteer-extra/tree/stealth-js
    # with open('./stealth.min.js') as f:
    #     browser.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
    #         "source": f.read()
    #     })

    # browser.get("https://bot.sannysoft.com/")
    # dummy_url = 'C:\Program Files (x86)'
    # browser.get('https://seller.shopee.tw')
    # browser.add_cookie({
    #     'name': '_gcl_au',
    #     'value': '1.1.1524688846.1662102542',
    #     'domain': 'seller.shopee.tw'
    # })
    # browser.add_cookie({
    #     'name': '_gid',
    #     'value': 'GA1.2.1769428283.1662102542',
    #     'domain': 'seller.shopee.tw'
    # })
    # browser.add_cookie({
    #     'name': 'shopee_webUnique_ccd',
    #     'value': 'sUaf3Zn2ueyMOYl8LOqkHQ%3D%3D%7C%2FUw0QDRC9P9mxWor8cpW4LOMBPoJw9F%2Bmli1rCx2lYjxumnxfGQHNIwqSCHZ5jg8ApIL3BxDCkZiPUDkRWJlZ05DpbQ%3D%7CO54JkZc74JxAEFKw%7C05%7C3',
    #     'domain': 'seller.shopee.tw'
    # })
    # browser.add_cookie({
    #     'name': '_ga',
    #     'value': 'GA1.2.777056962.1662102542',
    #     'domain': 'seller.shopee.tw'
    # })
    # browser.add_cookie({
    #     'name': '_QPWSDCXHZQA',
    #     'value': '2d6c8a8d-30b2-49b9-a8a0-42c35f2feea7',
    #     'domain': 'seller.shopee.tw'
    # })
    # browser.add_cookie({
    #     'name': '_fbp',
    #     'value': 'fb.1.1662102542452.1647776730',
    #     'domain': 'seller.shopee.tw'
    # })
    # browser.add_cookie({
    #     'name': 'SPC_CDS',
    #     'value': 'aaadbbd3-4df4-48c1-b176-eba3647f075e',
    #     'domain': 'seller.shopee.tw'
    # })
    # browser.add_cookie({
    #     'name': 'SPC_F',
    #     'value': 'CGvtARuzabSbElFagKIGjGGnNKPiaTfM',
    #     'domain': 'seller.shopee.tw'
    # })
    # browser.add_cookie({
    #     'name': 'SPC_STK',
    #     'value': '\"0h5WXl3ptLqoHUWs0CpolJdLnnDszja2YWUdHNwsth3lMNPiypzZIXWrsTh9pVJ8PSj6EVfR7l5bHvI2TZ3TbhlwnPyWiZBBrzvjqjFLaOI3Qxrd2DsLzikJeipvKgZfCpWOkGskalnvcY/eas0QqyoatqUAKKwovh6qrvpALkI=\"',
    #     'domain': 'seller.shopee.tw'
    # })
    # browser.add_cookie({
    #     'name': 'SPC_U',
    #     'value': '158512705',
    #     'domain': 'seller.shopee.tw'
    # })
    # browser.add_cookie({
    #     'name': 'SPC_R_T_ID',
    #     'value': '8QEAiEHRvMJe/cGVfu6BbR7SzVxBMZPXQbF/+GRUUmwuHzahf0K343W/ySSsLBMhYclbgrMitnvANIHZpGDwPnjoPtJLZjNxv52Merspe4ywmxGoZBfaLr7vkyjJpGNLEAkEAFy8mHb+N5v2kl8pX2K+SDc+VcWtOwyUoMyFaV8=',
    #     'domain': 'seller.shopee.tw'
    # })
    # browser.add_cookie({
    #     'name': 'SPC_R_T_IV',
    #     'value': 'NEJhZlgxUVVsUk13UmVxOA==',
    #     'domain': 'seller.shopee.tw'
    # })

    # 正玉佳
    # browser.add_cookie({
    #     'name': '_gcl_au',
    #     'value': '1.1.45462234.1621242076',
    #     'domain': 'seller.shopee.tw'
    # })
    # browser.add_cookie({
    #     'name': 'SPC_F',
    #     'value': '3FaDIdE1EpSzyqryjTUii4Fu7ZSKtotk',
    #     'domain': 'seller.shopee.tw'
    # })
    # browser.add_cookie({
    #     'name': '_ga',
    #     'value': 'GA1.1.1346315034.1621242077',
    #     'domain': 'seller.shopee.tw'
    # })
    # browser.add_cookie({
    #     'name': 'SPC_R_T_ID',
    #     'value': '\"LfRvt34krfN0wX1bXE3Aj6VjfM1PJsbPnp9O6DrUprhahKI0G7gRP8uBGwM6vt6A0iGa/yScB3zdkK4PIpMtUseJxx+yp0TG6xgvS7X/H9E=\"',
    #     'domain': 'seller.shopee.tw'
    # })
    # browser.add_cookie({
    #     'name': 'SPC_R_T_IV',
    #     'value': '\"csz04Yhq8zxy014o0QaxhA==\"',
    #     'domain': 'seller.shopee.tw'
    # })
    # browser.add_cookie({
    #     'name': '_ga_RPSBE3TQZZ',
    #     'value': 'GS1.1.1621242076.1.1.1621242225.60',
    #     'domain': 'seller.shopee.tw'
    # })
    # browser.add_cookie({
    #     'name': 'SPC_CDS',
    #     'value': 'a23ac5f1-f85c-4af1-9a75-e6cf580b7d34',
    #     'domain': 'seller.shopee.tw'
    # })
    # browser.add_cookie({
    #     'name': 'SPC_STK',
    #     'value': '\"3AL8lGlT+hwNQSxLKRZ9P+EudmGtvcQEJuFCRl8ih2/ZxyUP+DUqwqJeS7ZWhXhQXJ5AW1S5kyWTQFv70sRJ0oyCVs9s/dafhSpUfqgqt/w4eT/C3j9Y965OKMu+nY1Vz1Es/zcUKpnIKHwVZk9OUIUagajv9QSHJjZrqVDXE0Q=\"',
    #     'domain': 'seller.shopee.tw'
    # })
    # browser.add_cookie({
    #     'name': 'SPC_U',
    #     'value': '328318444',
    #     'domain': 'seller.shopee.tw'
    # })
    time.sleep(1)
    url = 'https://seller.shopee.tw/portal/sale/mass/ship'
    browser.get(url)
    time.sleep(1)
    browser.maximize_window()
    time.sleep(5)

    user_id = entry1.get()
    user_pwd = entry2.get()

    # 讀取資料庫獲取訂單基本資料(訂單號碼 訂單日期 客戶代號)
    # 646 改抓 Email1
    query_sql = "SELECT SunCust1M.CustNo, SunCust1M.TakeNo, SunCust1M.TakeNo2, SunCust1M.AccountWay, SunCust1M.Address, SunCust1M.Contact FROM SunCust1M WHERE ClassNo = '%s'" % user_id
    cursor.execute(query_sql)
    result = cursor.fetchone()
    if result is not None:
        customer_no = result[0]
        take_no = result[1]
        take_no2 = result[2]
        account_way = result[3]
        customer_address = result[4]
        customer_contact = result[5]
    else:
        customer_no = ''
        take_no = ''
        take_no2 = ''
        account_way = ''
        customer_address = ''
        customer_contact = ''

    # 3/24 蝦皮修改登入頁面
    while 'signin' in browser.current_url or 'verify' in browser.current_url or 'login?' in browser.current_url:
        #點選登入
        try:
            acc_elem = browser.find_element(By.XPATH, "//input[@name='loginKey']")
            print(acc_elem)
            if acc_elem and (user_id not in acc_elem.get_attribute("value")):
                acc_elem.clear()
                acc_elem.send_keys(user_id)
            pwd_elem = browser.find_element(By.XPATH, "//input[@name='password']")
            print(pwd_elem)
            if pwd_elem and (user_pwd not in pwd_elem.get_attribute("value")):
                pwd_elem.clear()
                pwd_elem.send_keys(user_pwd)
            time.sleep(1)
            input_elem = browser.find_elements(By.XPATH, "//input[@placeholder='驗證碼']")
            if len(input_elem) > 0 and (len(input_elem[0].get_attribute("value")) < 4):
                print('需要圖形驗證碼, 請協助輸入.')
                time.sleep(5)
            else:
                print('無驗證碼或已輸入驗證碼, 將點擊登入.')
                browser.find_element(By.XPATH, "//button[text()='登入']").click()
                # browser.find_elements(By.CLASS_NAME, 'wyhvVD')[1].click()
            time.sleep(3)
        except NoSuchElementException:
            # print('No login button detected. Login success.')
            # print(traceback.format_exc())
            pass

    time.sleep(12)
    if 'mass/ship' not in browser.current_url:
        elem = browser.find_elements(By.CLASS_NAME, 'sidebar-submenu-item-link')[1]
        elem.click()
        time.sleep(5)
    # 蝦皮新增功能概略
    while True:
        try:
            skip_btn = browser.find_element(By.CLASS_NAME, 'btn-skip')
            print('蝦皮新功能介紹，點選略過')
            skip_btn.click()
            time.sleep(1)
        except NoSuchElementException:
            print('已略過所有蝦皮新功能介紹，檢查蝦皮導覽')
            time.sleep(1)
            break
    # 蝦皮賣場導覽概略
    while True:
        try:
            skip_btn = browser.find_element(By.CLASS_NAME, 'shopee-guide-tooltip-button')
            print('蝦皮賣場導覽，點選略過')
            skip_btn.click()
            time.sleep(1)
        except NoSuchElementException:
            print('已略過所有蝦皮導覽，檢查蝦皮新介面介紹')
            time.sleep(1)
            break
    # 蝦皮新版介面概略
    while True:
        try:
            skip_btn = browser.find_element(By.CLASS_NAME, 'skip-button')
            print('蝦皮新介面介紹，點選略過')
            skip_btn.click()
            time.sleep(1)
        except NoSuchElementException:
            print('已略過所有蝦皮新介面介紹，開始正常掃描單號')
            time.sleep(5)
            break
    # 滑過新增的帳號資訊頁面
    element_to_hover_over = browser.find_element(By.CLASS_NAME, "subaccount-name")
    hover = ActionChains(browser).move_to_element(element_to_hover_over)
    hover.perform()
    element_to_hover_over = browser.find_element(By.CLASS_NAME, "eds-tabs__nav-tabs")
    hover = ActionChains(browser).move_to_element(element_to_hover_over)
    hover.perform()
    # 下載出貨文件
    elem = browser.find_elements(By.CLASS_NAME, 'tab-label')[1]
    elem.click()
    time.sleep(3)
    # 依照物流方式篩選訂單:
    filters = browser.find_element(By.CLASS_NAME, 'mass-ship-filter')
    filter_version = 'one'
    try:
        if (filters.find_element(By.CLASS_NAME, 'shipping-channel-filter')):
            filter_version = 'two'
            ship_list = filters.find_element(By.CLASS_NAME, 'shipping-channel-filter').find_elements(By.CLASS_NAME, 'shopee-radio-button__label')
            if len(ship_list) < 1:
                ship_list = filters.find_element(By.CLASS_NAME, 'shipping-channel-filter').find_elements(By.CLASS_NAME, 'eds-radio-button__label')
                print('eds version, get ship filter.')
                filter_version = 'eds'
            status_list = filters.find_element(By.CLASS_NAME, 'order-status-filter').find_elements(By.CLASS_NAME, 'shopee-radio-button__label')
            if len(status_list) > 2:
                status_list[1].click()
                time.sleep(1)
    except NoSuchElementException:
        try:
            ship_list = filters.find_elements(By.CLASS_NAME, 'eds-radio-button__label')
            if len(ship_list) < 1:
                print('Old version, get ship filter only.')
                ship_list = filters.find_elements(By.CLASS_NAME, 'shopee-radio-button__label')
            else:
                print('eds version, get ship filter only.')
                filter_version = 'eds'
        except NoSuchElementException:
            print('Old version, get ship filter only.')
            ship_list = filters.find_elements(By.CLASS_NAME, 'shopee-radio-button__label')
    problem_orders = []
    for shipping in ship_list:
        not_long_pre = True
        try:
            target = browser.find_element(By.CLASS_NAME, 'pre-order-filter')
            scrollElementIntoMiddle = "var viewPortHeight = Math.max(document.documentElement.clientHeight, window.innerHeight || 0);var elementTop = arguments[0].getBoundingClientRect().top;window.scrollBy(0, elementTop-(viewPortHeight/2));"
            browser.execute_script(scrollElementIntoMiddle, target)
            if filter_version == 'eds':
                preorder_filters = target.find_elements(By.CLASS_NAME, 'eds-radio-button__label')
            else:
                preorder_filters = target.find_elements(By.CLASS_NAME, 'shopee-radio-button__label')
            if len(preorder_filters) > 1:
                preorder_filters[1].click()
                print('Back to normal pre-order')
            time.sleep(2)
        except NoSuchElementException:
            print('No pre-order-filter')
        browser.find_element(By.TAG_NAME, 'html').send_keys(Keys.HOME)
        time.sleep(3)
        shipping.click()
        class_no = ''
        if product_no == '646':
            class_no = class_no_dict_646[shipping.text.split('(')[0].strip()]
        else:
            class_no = class_no_dict[shipping.text.split('(')[0].strip()]
        # print('解析: '+ class_no)
        time.sleep(4)
        no_package = False
        if class_no == '08' or (product_no == '646' and class_no == 'K'):
            print('其他寄送方式 - 檢查有無切換')
            if filter_version == 'two':
                print('其他寄送方式，切換至待處理訂單.')
                status_list = browser.find_element(By.CLASS_NAME, 'order-status-filter').find_elements(By.CLASS_NAME, 'shopee-radio-button__label')
                status_list[2].click()
                time.sleep(2)
            elif filter_version == 'one':
                shopee_selectors = browser.find_elements(By.CLASS_NAME, 'shopee-selector__inner')
                for shopee_selector in shopee_selectors:
                    if shopee_selector.text.strip() == '已處理':
                        print('其他寄送方式需切換至全部訂單.')
                        shopee_selector.click()
                        time.sleep(2)
                        shopee_options = browser.find_elements(By.CLASS_NAME, 'shopee-option')
                        for shopee_option in shopee_options:
                            if shopee_option.text.strip() == '全部訂單 出貨狀態':
                                browser.execute_script("arguments[0].click();", shopee_option)
                                # shopee_option.click()
                                print('切換至全部訂單')
                                no_package = True
                                time.sleep(2)
                                break
                        else:
                            continue
                        break
            elif filter_version == 'eds':
                print('其他寄送方式需切換至全部訂單.(eds)')
                try:
                    shopee_selectors = browser.find_elements(By.CLASS_NAME, 'order-list-filter-item')
                    shopee_selectors[2].click()
                    time.sleep(1)
                    eds_options = browser.find_elements(By.CLASS_NAME, 'eds-option')
                    for eds_option in eds_options:
                        if eds_option.text.strip() == '全部訂單 出貨狀態':
                            browser.execute_script("arguments[0].click();", eds_option)
                            # eds_option.click()
                            print('切換至全部訂單')
                            no_package = True
                            time.sleep(2)
                            break
                        else:
                            continue
                except Exception as e:
                    print(e)
        if filter_version == 'two':
            try:
                pages = browser.find_element(By.CLASS_NAME, 'shopee-pager__pages')
                current_page = int(pages.find_element(By.CLASS_NAME, 'active').text)
                if pages.find_elements(By.CLASS_NAME, 'shopee-pager__page')[-1].text == '...':
                    total_page = int(pages.find_elements(By.CLASS_NAME, 'shopee-pager__page')[-2].text)
                elif len(pages.find_elements(By.CLASS_NAME, 'shopee-pager__page')) > 1:
                    total_page = int(pages.find_elements(By.CLASS_NAME, 'shopee-pager__page')[-1].text)
                else:
                    total_page = int(pages.find_elements(By.CLASS_NAME, 'shopee-pager__page')[0].text)
            except NoSuchElementException:
                print('No pages. Next shipping.')
                continue
        elif filter_version == 'one':
            current_page = int(browser.find_element(By.CLASS_NAME, 'shopee-pager__current').text)
            total_page = int(browser.find_element(By.CLASS_NAME, 'shopee-pager__total').text)
        elif filter_version == 'eds':
            try:
                pages = browser.find_element(By.CLASS_NAME, 'eds-pager__pages')
                eds_pages = pages.find_elements(By.CLASS_NAME, 'eds-pager__page')
                if len(eds_pages) > 0:
                    current_page = int(pages.find_element(By.CLASS_NAME, 'active').text)
                    if eds_pages[-1].text == '...':
                        total_page = int(eds_pages[-2].text)
                    elif len(eds_pages) > 1:
                        total_page = int(eds_pages[-1].text)
                    else:
                        total_page = int(eds_pages[0].text)
                else:
                    current_page = int(browser.find_element(By.CLASS_NAME, 'eds-pager__current').text)
                    total_page = int(browser.find_element(By.CLASS_NAME, 'eds-pager__total').text)
            except NoSuchElementException:
                print('No pages. Next shipping.')
                continue
        while current_page <= total_page:
            time.sleep(3)
            shopee_table_rows = browser.find_elements(By.CLASS_NAME, 'mass-ship-row-wrapper')
            for row in shopee_table_rows:
                old_page = browser.window_handles[0]
                target = row.find_element(By.CLASS_NAME, 'order-id')
                scrollElementIntoMiddle = "var viewPortHeight = Math.max(document.documentElement.clientHeight, window.innerHeight || 0);var elementTop = arguments[0].getBoundingClientRect().top;window.scrollBy(0, elementTop-(viewPortHeight/2));"
                browser.execute_script(scrollElementIntoMiddle, target)
                quot_customer = target.text
                if (len(quot_customer.split('\n')) > 1):
                    quot_customer = quot_customer.split('\n')[1]
                print('解析: '+ quot_customer)
                if product_no == '078' or product_no == '097': # 正玉佳金固開啟OK蝦皮互通 需再判斷物流是不是OK 合日昇沒使用蝦皮店到店 預設為OK
                    target_shipping = row.find_element(By.CLASS_NAME, 'shipping-option').text.strip()
                    if 'OK' in target_shipping:
                        class_no = '04'
                    print('解析: '+ class_no)
                # 檢查客戶訂號 如已有則不重複建檔
                check_quot_query_sql = "SELECT SaleNo FROM SunSale1M WHERE OrdrCust = '%s'" % quot_customer
                cursor.execute(check_quot_query_sql)
                result = cursor.fetchone()
                # print(result)
                order_no = now.strftime('%Y%m%d') + '0001'
                if result is not None:
                    print(quot_customer + ' - Order created, skip to next order.')
                    new_order = False
                    continue
                else:
                    order_no_query_sql = "SELECT MAX(SaleNo) FROM SunSale1M WHERE SaleNo LIKE '{}%'".format(now.strftime('%Y%m%d'))
                    cursor.execute(order_no_query_sql)
                    result = cursor.fetchone()
                    if result[0] is not None:
                        if (int(order_no) <= int(result[0])):
                            order_no = str(int(result[0]) + 1)
                print('before click')
                print(target)
                link = target.find_element(By.TAG_NAME, 'a')
                link.click()
                print('after click')
                print(target)
                new_order = True
                while new_order == True:
                    try:
                        time.sleep(3)
                        new_page = browser.window_handles[1]
                        browser.switch_to.window(new_page)
                        time.sleep(2)
                        if no_package:
                            order_payment = ''
                        else:
                            package_info = browser.find_element(By.CLASS_NAME, 'log-info-header')
                            if filter_version == 'two':
                                try:
                                    package_name = package_info.find_element(By.CLASS_NAME, 'carrier').text + package_info.find_element(By.CLASS_NAME, 'label').text
                                except NoSuchElementException:
                                    print('No package number. Record carrier only')
                                    package_name = package_info.find_element(By.CLASS_NAME, 'carrier').text
                            else:
                                package_name = package_info.find_element(By.CLASS_NAME, 'label').text
                            order_payment = package_name.replace(" ","")
                        # print('解析: '+ order_payment)
                        item_list = []
                        # item list
                        product_list = browser.find_elements(By.CLASS_NAME, 'product-list-item')
                        for product in product_list[1:]:
                            product_meta_split = product.find_element(By.CLASS_NAME, 'product-meta').text.split('商品選項貨號:')
                            # 合日昇 箱購產品使用 quan1
                            product_name = product.find_element(By.CLASS_NAME, 'product-name').text
                            use_quan1 = False
                            item_unit = ''
                            if '箱購' in product_name:
                                use_quan1 = True
                            if len(product_meta_split) > 1:
                                item_no = product_meta_split[1].strip()
                            else:
                                if product_no == '078':
                                    item_no = '0'
                                else:
                                    item_no = 'ZZZZZZZ1'
                                problem_order = (order_no, quot_customer, '沒有填寫商品貨號')
                                problem_orders.append(problem_order)
                            price = product.find_element(By.CLASS_NAME, 'price').text.replace(",","").replace(" ", "")
                            quan2 = product.find_element(By.CLASS_NAME, 'qty').text.replace(",","").replace(" ", "")
                            check_bundle = False
                            item_percent = 1.0
                            item_total = 0
                            item_change = 1
                            try:
                                if (product.find_element(By.CLASS_NAME, 'bundle-price')):
                                    check_bundle = True
                                    print('Got bundle price, bundle price used')
                            except NoSuchElementException:
                                print('No bundle price, single price used')
                            if check_bundle:
                                item_total = product.find_element(By.CLASS_NAME, 'price-before-bundle').text.replace(",","").replace(" ", "")
                                item_percent = int(product.find_element(By.CLASS_NAME, 'price-after-bundle').text.replace(",","").replace(" ", ""))/int(product.find_element(By.CLASS_NAME, 'price-before-bundle').text.replace(",","").replace(" ", ""))
                            else:
                                item_total = product.find_element(By.CLASS_NAME, 'subtotal').text.replace(",","")
                            # 正玉佳使用國際條碼
                            if product_no == '078':
                                if not item_no.startswith('('):
                                    item_no = item_no.split('(')[0]
                                check_item_no = item_no.split(')')
                                if len(check_item_no) > 1:
                                    item_no = check_item_no[1]
                                # 2025/04/01 李依霖Mia 條碼抓取到多個單位，選用Ord較小的單位
                                query_product_sql = "SELECT ItemNo, Unit, Change FROM SunItemUnit WHERE BarCode = '%s' ORDER BY Ord" % item_no
                                cursor.execute(query_product_sql)
                                result = cursor.fetchone()
                                # print(result)
                                if result is not None:
                                    item_no = result[0]
                                    item_unit = result[1]
                                    item_change = result[2]
                                else:
                                    query_product_sql = "SELECT ItemNo, Unit, Change FROM SunItemUnit WHERE ItemNo = '%s' ORDER BY Ord" % item_no
                                    cursor.execute(query_product_sql)
                                    result = cursor.fetchone()
                                    # print(result)
                                    if result is not None:
                                        item_no = result[0]
                                        item_unit = result[1]
                                        item_change = result[2]
                                    else:
                                        item_no = '0'
                                        problem_order = (order_no, quot_customer, '沒有填寫商品貨號')
                                        problem_orders.append(problem_order)
                            item_list.append((item_no, price, quan2, item_total, item_percent, use_quan1, item_unit, item_change))
                        # bundle item list
                        bundle_list = browser.find_elements(By.CLASS_NAME, 'order-payment-item')
                        for bundle in bundle_list:
                            bundle_product_list = bundle.find_elements(By.CLASS_NAME, 'bundle-deal-item')
                            for bproduct in bundle_product_list:
                                product_meta_split = bproduct.find_element(By.CLASS_NAME, 'product-meta').text.split('商品選項貨號:')
                                # 合日昇 箱購產品使用 quan1
                                product_name = bproduct.find_element(By.CLASS_NAME, 'product-name').text
                                use_quan1 = False
                                item_unit = ''
                                if '箱購' in product_name:
                                    use_quan1 = True
                                if len(product_meta_split) > 1:
                                    item_no = product_meta_split[1].strip()
                                elif product_no == '097':
                                    item_no = 'ZZZZZZZ1'
                                    problem_order = (order_no, quot_customer, '沒有填寫商品貨號')
                                    problem_orders.append(problem_order)
                                elif product_no == '078':
                                    item_no = '0'
                                    problem_order = (order_no, quot_customer, '沒有填寫商品貨號')
                                    problem_orders.append(problem_order)
                                else:
                                    item_no = ''
                                    problem_order = (order_no, quot_customer, '沒有填寫商品貨號')
                                    problem_orders.append(problem_order)
                                price = bproduct.find_element(By.CLASS_NAME, 'price').text.replace(",","").replace(" ", "")
                                quan2 = bproduct.find_element(By.CLASS_NAME, 'qty').text.replace(",","").replace(" ", "")
                                check_bundle = False
                                item_percent = 1.0
                                item_total = 0
                                item_change = 1
                                try:
                                    if (bundle.find_element(By.CLASS_NAME, 'bundle-price')):
                                        check_bundle = True
                                        print('Got bundle price, bundle price used')
                                except NoSuchElementException:
                                    print('No bundle price, single price used')
                                if check_bundle:
                                    item_total = bundle.find_element(By.CLASS_NAME, 'price-before-bundle').text.replace(",","").replace(" ", "")
                                    item_percent = int(bundle.find_element(By.CLASS_NAME, 'price-after-bundle').text.replace(",","").replace(" ", ""))/int(bundle.find_element(By.CLASS_NAME, 'price-before-bundle').text.replace(",","").replace(" ", ""))
                                else:
                                    item_total = bundle.find_element(By.CLASS_NAME, 'subtotal').text.replace(",","")
                                # 正玉佳和合日昇使用國際條碼
                                if product_no == '078' or product_no == '646':
                                    if not item_no.startswith('('):
                                        item_no = item_no.split('(')[0]
                                    check_item_no = item_no.split(')')
                                    if len(check_item_no) > 1:
                                        item_no = check_item_no[1]
                                    query_product_sql = "SELECT ItemNo, Unit, Change FROM SunItemUnit WHERE BarCode = '%s' ORDER BY Ord" % item_no
                                    cursor.execute(query_product_sql)
                                    result = cursor.fetchone()
                                    # print(result)
                                    if result is not None:
                                        item_no = result[0]
                                        item_unit = result[1]
                                        item_change = result[2]
                                    else:
                                        if product_no == '078':
                                            item_no = '0'
                                        else:
                                            item_no = 'ZZZZZZZ1'
                                        problem_order = (order_no, quot_customer, '沒有填寫商品貨號')
                                        problem_orders.append(problem_order)
                                item_list.append((item_no, price, quan2, item_total, item_percent, use_quan1, item_unit, item_change))
                        # print('解析: ')
                        # print(item_list)
                        income_container = browser.find_element(By.CLASS_NAME, 'income-container')
                        price_tags = income_container.find_elements(By.CLASS_NAME, 'income-label')
                        no_fee = True
                        has_coupon = False
                        for tag in price_tags:
                            if tag.text == '折扣券' or tag.text == '優惠券':
                                has_coupon = True
                            if '手續費' in tag.text:
                                no_fee = False
                        total_prices = browser.find_elements(By.CLASS_NAME, 'income-subtotal')
                        for _ in range(5):
                            if len(total_prices) > 1:
                                print('price get success.')
                                break
                            else:
                                print('price get fail. Retry...')
                                time.sleep(1)
                                total_prices = browser.find_elements(By.CLASS_NAME, 'income-subtotal')
                        # 蝦皮改版 更改抓取價格位置(2/11)
                        # 蝦皮改版 更改抓取價格位置(6/28)
                        # 蝦皮改版 更改抓取價格位置(11/4)
                        # print(total_prices)
                        total_price = int(total_prices[0].find_element(By.CLASS_NAME, 'income-value').text.replace("NT$", "").replace(",", "").replace("$", "").replace(" ", ""))
                        shipping_price = int(total_prices[1].find_element(By.CLASS_NAME, 'income-value').text.replace("NT$", "").replace(",", "").replace("$", "").replace(" ", ""))
                        order_fee = int(total_prices[2].find_element(By.CLASS_NAME, 'income-value').text.replace("NT$", "").replace(",", "").replace("$", "").replace(" ", ""))
                        coupon_price = 0
                        if has_coupon:
                            coupon_price = int(total_prices[2].find_element(By.CLASS_NAME, 'income-value').text.replace("NT$", "").replace(",", "").replace("$", "").replace(" ", ""))
                            order_fee = int(total_prices[3].find_element(By.CLASS_NAME, 'income-value').text.replace("NT$", "").replace(",", "").replace("$", "").replace(" ", ""))
                        # print('解析: '+ str(total_price))
                        # print('解析: '+ str(shipping_price))
                        # print('解析: '+ str(coupon_price))
                        # print('解析: '+ str(order_fee))
                        # 總價折扣(手續費折扣券攤入各商品)
                        # coupon_price and order_fee should be minus
                        if class_no == '08' or (product_no == '646' and class_no == 'K'):
                            final_price = total_price + shipping_price + coupon_price + order_fee
                        else:
                            final_price = total_price + coupon_price + order_fee
                        # print('解析: '+ str(final_price))
                        shopee_discount = (total_price + order_fee) / total_price
                        tax_no = 0
                        # 正玉佳 稅內含
                        if product_no == '078':
                            tax_no = 1
                        # print('解析: '+ str(shopee_discount))
                        # 寫入至database
                        order_id = str(uuid.uuid4()).upper()
                        order_no_query_sql = "SELECT MAX(SaleNo) FROM SunSale1M WHERE SaleNo LIKE '{}%'".format(now.strftime('%Y%m%d'))
                        cursor.execute(order_no_query_sql)
                        result = cursor.fetchone()
                        if result[0] is not None:
                            if (int(order_no) <= int(result[0])):
                                order_no = str(int(result[0]) + 1)
                        save_order_query = "INSERT INTO SunSale1M " \
                            "(SaleID, StorNo, SaleNo, SaleDate, CompNo, CustNo, CustNo2, CustAddr, CustContact, " \
                            "Accountway, AccMonth, TaxNo, TaxRate, DollNo, DollRate, IsNation, " \
                            "TakeNo1, TakeNo2, Locked, Memo2, EffectiveDate, BuildDate, BuildUser, EditDate, EditUser, CustNo3, " \
                            "IsNotTran, IsTran, Chk1, Chk2, SPChk, OrdrCust, TaxCalc) " \
                            "VALUES ('{ordrid}', '01', '{ordrno}', '{ordrdate}', '01', '{custno}', '{custno}', '{custaddr}', '{custcontact}', " \
                            "'{accountway}', '{accmonth}', {taxno}, 0.05, 'NTD', 1, 1, " \
                            "'{takeno1}', '{takeno2}', 0, '{memo2}', '{effectivedate}', '{builddate}', '00', '{editdate}', '00', '{custno}', " \
                            "0, 0, 0, 0, 0, '{quot_customer}', 0);".format(
                                ordrid=order_id,
                                ordrno=order_no,
                                ordrdate=order_date,
                                custno=customer_no,
                                custaddr=customer_address,
                                custcontact=customer_contact,
                                accountway=account_way,
                                accmonth=account_year,
                                taxno=tax_no,
                                takeno1=take_no,
                                takeno2=take_no2,
                                memo2=memo_string,
                                effectivedate=order_date,
                                builddate=build_date,
                                editdate=order_date,
                                quot_customer=quot_customer,
                            )
                        # print(save_order_query)
                        cursor.execute(save_order_query)
                        conn.commit()
                        item_ord = 1
                        # 商品品項
                        for item in item_list:
                            # item tuple: (item_no, price, quan2, item_total, item_percent, use_quan1, item_unit, item_change)
                            sql = "SELECT Unit, Change \
                            FROM SunItemUnit \
                            WHERE ItemNo = '%s' ORDER BY Ord" % (item[0])
                            cursor.execute(sql,)
                            result = cursor.fetchone()
                            unit = ''
                            change = 1
                            if result is not None:
                                unit = result[0]
                                change = result[1]
                            if item[6]:
                                unit = item[6]
                                change = item[7]
                            item_discount = round(item[4] * shopee_discount, 3)
                            # item_quan1 = 0
                            # item_quan2 = item[2]
                            # if check_unit:
                            item_quan1 = item[2]
                            item_quan2 = 0
                            order_ord = str(uuid.uuid4()).upper()
                            save_order_item_query = "INSERT INTO SunSale1D " \
                                "(SaleOrd, SaleID, Ord, Type1, Type2, Type3, " \
                                "ItemNo, Unit, UnitQuan, UnitPrice, UnitMoney, " \
                                "Quan, HaveQuan, Discount, CondNo, Remark, ProjID, WhichNo, PDAID, " \
                                "Allowance, PromID, StockRemark, CustProjID, CostPriceN) " \
                                "VALUES ('{ordrord}', '{ordrid}', {ord}, 0, 0, 0, " \
                                "'{itemno}', '{unit}', {unitquan}, {unitprice}, {unitmoney}, " \
                                "{quan}, 0.0, {discount}, NULL, '{remark}', NULL, '01', NULL, " \
                                "NULL, NULL, NULL, NULL, NULL);".format(
                                    ordrord=order_ord,
                                    ordrid=order_id,
                                    ord=item_ord,
                                    itemno=item[0],
                                    unit=unit,
                                    unitquan=item_quan1,
                                    unitprice=item[1],
                                    unitmoney=round(int(item[1]) * int(item[2]) * item_discount, subtotal_places),
                                    quan=int(item_quan1) * int(change),
                                    discount=item_discount,
                                    remark=quot_customer,
                                )
                            # print(save_order_item_query)
                            cursor.execute(save_order_item_query)
                            conn.commit()
                            item_ord = item_ord + 1
                        # 運費品項
                        shipping_item_no = '9999999999'
                        # 正玉佳 有運費品項
                        if product_no == '078':
                            shipping_item_no = '00'
                        if (class_no == '08' or (product_no == '646' and class_no == 'K')) and shipping_price != 0:
                            order_ord = str(uuid.uuid4()).upper()
                            save_order_item_query = "INSERT INTO SunSale1D " \
                                "(SaleOrd, SaleID, Ord, Type1, Type2, Type3, " \
                                "ItemNo, Unit, UnitQuan, UnitPrice, UnitMoney, " \
                                "Quan, HaveQuan, Discount, CondNo, Remark, ProjID, WhichNo, PDAID, " \
                                "Allowance, PromID, StockRemark, CustProjID, CostPriceN) " \
                                "VALUES ('{ordrord}', '{ordrid}', {ord}, 0, 0, 0, " \
                                "'{itemno}', '{unit}', 1, {unitprice}, {unitmoney}, " \
                                "1, 0.0, 1.0, NULL, '{remark}', NULL, '01', NULL, " \
                                "NULL, NULL, NULL, NULL, NULL);".format(
                                    ordrord=order_ord,
                                    ordrid=order_id,
                                    ord=item_ord,
                                    itemno=shipping_item_no,
                                    unit=unit,
                                    unitprice=shipping_price,
                                    unitmoney=shipping_price,
                                    remark=quot_customer,
                                )
                            print('SHIPPING FEE USED')
                            # print(save_order_item_query)
                            cursor.execute(save_order_item_query)
                            conn.commit()
                            item_ord = item_ord + 1
                        # 折扣券品項
                        if has_coupon:
                            order_ord = str(uuid.uuid4()).upper()
                            save_order_item_query = "INSERT INTO SunSale1D " \
                                "(SaleOrd, SaleID, Ord, Type1, Type2, Type3, " \
                                "ItemNo, Unit, UnitQuan, UnitPrice, UnitMoney, " \
                                "Quan, HaveQuan, Discount, CondNo, Remark, ProjID, WhichNo, PDAID, " \
                                "Allowance, PromID, StockRemark, CustProjID, CostPriceN) " \
                                "VALUES ('{ordrord}', '{ordrid}', {ord}, 0, 0, 0, " \
                                "'999999999', '{unit}', 1, {unitprice}, {unitmoney}, " \
                                "1, 0.0, 1.0, NULL, '{remark}', NULL, '01', NULL, " \
                                "NULL, NULL, NULL, NULL, NULL);".format(
                                    ordrord=order_ord,
                                    ordrid=order_id,
                                    ord=item_ord,
                                    unit=unit,
                                    unitprice=coupon_price,
                                    unitmoney=coupon_price,
                                    remark=quot_customer,
                                )
                            print('COUPON USED')
                            # print(save_order_item_query)
                            cursor.execute(save_order_item_query)
                            conn.commit()
                            item_ord = item_ord + 1
                        print('Order {} saved.'.format(order_no))
                        # get shipping details
                        # checkbox = row.find_element(By.XPATH, "//input[@type='checkbox']")
                        # if not checkbox.is_selected():
                        #     checkbox.click()
                        # shipping_button = browser.find_element(By.CLASS_NAME, 'shopee-button--large')
                        # shipping_button.click()
                        # time.sleep(5)
                        # new_page = browser.window_handles[1]
                        # browser.switch_to.window(new_page)
                        # # 7-11
                        # if class_no == '01':
                        #     seven = browser.find_element(By.XPATH, "//button[@value='顯示交貨便服務單']")
                        #     seven.click()
                        #     time.sleep(1)
                        #     barcode = browser.find_element_by_id('lblBarCode1')
                        #     print('解析: ' + barcode.text)
                        #     barcode2 = browser.find_element_by_id('lblBarCode2')
                        #     print('解析: ' + barcode2.text)
                        #     barcode3 = browser.find_element_by_id('lblBarCode3')
                        #     print('解析: ' + barcode3.text)
                        #     pincode = browser.find_element_by_id('lblC2BPinCode1')
                        #     print('解析: ' + pincode.text)
                        #     recshop = browser.find_element_by_id('lblRecvshop1')
                        #     print('解析: ' + recshop.text)
                        #     receiver = browser.find_element_by_id('lblRecver1')
                        #     print('解析: ' + receiver.text)
                        #     barcode4 = browser.find_element_by_id('lblBarCode4')
                        #     print('解析: ' + barcode4.text)
                        #     deadline = browser.find_element_by_id('lblDeadDateTime1')
                        #     print('解析: ' + deadline.text)
                        #     paymentname = browser.find_element_by_id('lblPaymentCPName1')
                        #     print('解析: ' + paymentname.text)
                        #     orderno = browser.find_element_by_id('lblOrderNo')
                        #     print('解析: ' + orderno.text)
                        #     shopper = browser.find_element_by_id('lblShopperName1')
                        #     print('解析: ' + shopper.text)
                        # elif class_no == '02':
                        #     browser.find_element(By.CLASS_NAME, '')
                    except (NoSuchElementException, StaleElementReferenceException) as e:
                        # print('No login button detected. Login success.')
                        # print(traceback.format_exc())
                        pass
                    new_order = False
                    browser.close()
                    browser.switch_to.window(old_page)
            # next page of orders
            if current_page == total_page:
                try:
                    if not_long_pre:
                        target = browser.find_element(By.CLASS_NAME, 'pre-order-filter')
                        scrollElementIntoMiddle = "var viewPortHeight = Math.max(document.documentElement.clientHeight, window.innerHeight || 0);var elementTop = arguments[0].getBoundingClientRect().top;window.scrollBy(0, elementTop-(viewPortHeight/2));"
                        browser.execute_script(scrollElementIntoMiddle, target)
                        if filter_version == 'eds':
                            preorder_filters = target.find_elements(By.CLASS_NAME, 'eds-radio-button__label')
                        else:
                            preorder_filters = target.find_elements(By.CLASS_NAME, 'shopee-radio-button__label')
                        if len(preorder_filters) > 1:
                            preorder_filters[2].click()
                            not_long_pre = False
                            time.sleep(2)
                            continue
                    else:
                        print('pre-order-filter long-pre-pressed, Next shipping.')
                        break
                except NoSuchElementException:
                    print('No pre-order-filter Next shipping.')
                    break
            browser.find_element(By.TAG_NAME, 'html').send_keys(Keys.HOME)
            time.sleep(3)
            if filter_version == 'eds':
                next_page_button = browser.find_element(By.CLASS_NAME, 'eds-pager__button-next')
            else:
                next_page_button = browser.find_element(By.CLASS_NAME, 'shopee-pager__button-next')
            try:
                browser.execute_script("arguments[0].click();", next_page_button)
            except Exception as err:
                print(err)
            time.sleep(5)
            if filter_version == 'two':
                pages = browser.find_element(By.CLASS_NAME, 'shopee-pager__pages')
                current_page = int(pages.find_element(By.CLASS_NAME, 'active').text)
                if pages.find_elements(By.CLASS_NAME, 'shopee-pager__page')[-1].text == '...':
                    total_page = int(pages.find_elements(By.CLASS_NAME, 'shopee-pager__page')[-2].text)
                elif len(pages.find_elements(By.CLASS_NAME, 'shopee-pager__page')) > 1:
                    total_page = int(pages.find_elements(By.CLASS_NAME, 'shopee-pager__page')[-1].text)
                else:
                    total_page = int(pages.find_elements(By.CLASS_NAME, 'shopee-pager__page')[0].text)
            elif filter_version == 'one':
                current_page = int(browser.find_element(By.CLASS_NAME, 'shopee-pager__current').text)
                total_page = int(browser.find_element(By.CLASS_NAME, 'shopee-pager__total').text)
            elif filter_version == 'eds':
                pages = browser.find_element(By.CLASS_NAME, 'eds-pager__pages')
                eds_pages = pages.find_elements(By.CLASS_NAME, 'eds-pager__page')
                if len(eds_pages) > 0:
                    current_page = int(pages.find_element(By.CLASS_NAME, 'active').text)
                    if eds_pages[-1].text == '...':
                        total_page = int(eds_pages[-2].text)
                    elif len(eds_pages) > 1:
                        total_page = int(eds_pages[-1].text)
                    else:
                        total_page = int(eds_pages[0].text)
                else:
                    current_page = int(browser.find_element(By.CLASS_NAME, 'eds-pager__current').text)
                    total_page = int(browser.find_element(By.CLASS_NAME, 'eds-pager__total').text)
    # 關閉瀏覽器
    browser.quit()
    info_message = "完成訂單建檔, 有疑慮訂單如下:"
    for order_data in problem_orders:
        info_message += "\n訂單編號: {}, 客戶訂號: {}, 問題原因: {}".format(order_data[0], order_data[1], order_data[2])
    messagebox.showinfo("通知", info_message)

# GUI
# 建立視窗
window=Tk()
window.title("蝦皮批次出貨轉早陽XE訂單 %s 版" % ver_num)
# 視窗大小和位置(中間為字母x)
window.geometry('800x520')
window.geometry('+100+100')
# window.geometry("380x420+500+240")

# 標籤控制元件
label=Label(window,text='程式會自動開啟Chrome，請勿在自動開啟的Chrome上進行操作以免影響自動執行的結果',font=("微軟正黑體",15),fg="red")
# 定位 grid網格佈局 pack包 place位置
label.grid(row=0, column=0, columnspan=7)

label5=Label(window,text="登入蝦皮網頁帳號：", font=("微軟正黑體",16))
label5.grid(row=2, column=0)
entry1=Entry(window,font=("微軟正黑體",16))
entry1.insert(END, user_id)
entry1.grid(row=2, column=1)
label6=Label(window,text="登入蝦皮網頁密碼：", font=("微軟正黑體",16))
label6.grid(row=3, column=0)
entry2=Entry(window,font=("微軟正黑體",16))
entry2.insert(END, user_pwd)
entry2.grid(row=3, column=1)

# 按鈕控制元件
but=Button(window, text="執行", command=CalItem, font=("微軟正黑體",15))
but.grid(row=5, column=1, sticky=EW)

# LINE好友圖檔
img=Image.open("早陽QRCODE.jpg") 
img = img.resize((250, 250), Image.ANTIALIAS)
photo=ImageTk.PhotoImage(img)
label7=Label(image=photo)
label7.grid(row=8, column=4, columnspan=3)

label10=Label(window,text='LINE行動條碼：', font=("微軟正黑體",15))
label10.grid(row=7, column=4, columnspan=3)

# 標籤控制元件
label8=Label(window,text='● 進銷存管理系統      ● 會計系統  ● 專案設計\n● 行動商務系統          ● 薪資系統  ● POS系統\n● PDA庫存盤點系統  ● WEB網頁線上下單系統', font=("微軟正黑體",12), justify=LEFT, anchor="e")
# 定位 grid網格佈局 pack包 place位置
label8.grid(row=7, column=0, columnspan=3, sticky=NW)

# label11=Label(window,text='[聯絡我們]', font=("微軟正黑體",15))
# label11.grid(row=8, column=0, columnspan=3, sticky=W)

label9=Label(window,text='[聯絡我們]\n服務時間：星期一~星期五 早上9:00~下午6:00\n電話服務：06-235-2945\n服務信箱： sunrise@ms1.hinet.net\nLINE ID：@155fixeh', font=("微軟正黑體",15), justify=LEFT, anchor="w")
label9.grid(row=8, column=0, columnspan=3, sticky=NW)
window.grid_rowconfigure(1, minsize=4)
window.grid_rowconfigure(4, minsize=10)
window.grid_rowconfigure(6, minsize=45)

# 顯示視窗(訊息迴圈)
window.mainloop()
