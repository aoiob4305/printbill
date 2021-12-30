#-*- coding: utf-8 -*-

from msedge.selenium_tools import Edge, EdgeOptions
from selenium.webdriver.support.ui import Select

import csv
from time import sleep
from configparser import ConfigParser

from auto_utils import clickImgButton, fullpageScreenshot, pngToPdf, sendKey


def readCustomer(filename):
    customer_list = []
    #with open(filename, encoding='utf-8') as csvfile:
    with open(filename, encoding='cp949') as csvfile:
        reader = csv.reader(csvfile)
        for row in reader:
            customer = (row[0].strip(), row[1].strip(), row[2].strip().split('-'))
            customer_list.append(customer)

    return customer_list

def is_text_present(driver, text):
    return str(text) in driver.page_source

def openUrlByEdge(url):
    opts = EdgeOptions()
    opts.use_chromium = True

    driver = Edge(executable_path='msedgedriver', options=opts)
    driver.get(url)
    driver.implicitly_wait(time_to_wait=5)

    return driver

def closeAlert(driver):
    try:
        driver.switch_to.alert.accept()
    except:
        print("경고창이 없습니다.")


def loginKepco(driver, id, password):
    login_link = driver.find_element_by_link_text("로그인")
    login_link.click()

    login_idbox = driver.find_element_by_name("id_A")
    login_passbox = driver.find_element_by_name("pw")
    login_button = driver.find_element_by_class_name("btn_bbs")

    login_idbox.send_keys(id)
    login_passbox.send_keys(password)
    login_button.click()


def printBill(driver):
    #조회된 청구서 클릭
    driver.switch_to.window(driver.window_handles[1])
    driver.find_element_by_class_name("hovering").click()

    #인쇄, 팝업창뜨기까지 기다림
    sleep(5)

    driver.find_element_by_link_text("인쇄").click()
    sleep(10)

    driver.switch_to.window(driver.window_handles[1])
    sleep(10)

    clickImgButton("./printbutton.png")
    sleep(10)

    driver.switch_to.window(driver.window_handles[1])
    driver.close()
    driver.switch_to.window(driver.window_handles[0])    

def printBillToPdf(driver, filename):
    #조회된 청구서 클릭
    print("청구서를 클릭합니다.")
    driver.find_element_by_class_name("hovering").click()
    sleep(5)
    
    #인쇄, 팝업창뜨기까지 기다림
    print("인쇄버튼을 누릅니다.")
    driver.find_element_by_link_text("인쇄").click()
    sleep(10)
    
    print("다른창으로 전환합니다.")
    driver.switch_to.window(driver.window_handles[1])
    sleep(1)

    print("취소버튼을 누릅니다.")
    sendKey('esc')
    sleep(1)

    #파일 저장
    driver.switch_to.window(driver.window_handles[1])
    print("화면을 캡쳐하여 저장합니다.")
    fullpageScreenshot(driver, "temp.png")
    pngToPdf("temp.png", filename)
    
    sleep(5)

    driver.close()
    driver.switch_to.window(driver.window_handles[0])    

if __name__ == '__main__':
    
    config = ConfigParser()
    config.read('config.txt')

    URL = config['kepco']['URL']
    USER_ID = config['user']['USER_ID']
    USER_PASS = config['user']['USER_PASS']

    CUSTOMER = readCustomer('customer.csv')

    #메인화면
    driver = openUrlByEdge(URL)
    closeAlert(driver)

    loginKepco(driver, USER_ID, USER_PASS)
    closeAlert(driver) #로그인 환영메시지창 넘기기

    #조회화면까지들어가기
    driver.find_element_by_partial_link_text("납부").click()

    #개인보호 안내강화창 넘기기
    closeAlert(driver)

    for idx, customer_num in enumerate(CUSTOMER):
        #상세조회로 이동
        driver.find_element_by_link_text("상세조회").click()

        #상세조회 이용안내 팝업 넘기기
        driver.find_element_by_class_name("close_box").click()

        #검색조건 선택/입력 및 조회
        Select(driver.find_element_by_id("sch")).select_by_visible_text("고객번호")
        customer_num_box = driver.find_element_by_name("cust_num")
        customer_num_box.send_keys(customer_num[0])
        driver.find_element_by_link_text("조회").click()

        #고객정보 확인창 입력
        sleep(2)
        driver.switch_to.window(driver.window_handles[1])
        if customer_num[1] == "법인고객":
            selected = Select(driver.find_element_by_name("jumin_select"))
            selected.select_by_visible_text(customer_num[1])
            print(customer_num[2])
            driver.find_element_by_name("jumin_tmp_03").send_keys(customer_num[2][0])
            driver.find_element_by_name("jumin_tmp_04").send_keys(customer_num[2][1])
        elif customer_num[1] == "국가기관(사업자)":
            selected = Select(driver.find_element_by_name("jumin_select"))
            selected.select_by_visible_text(customer_num[1])
            driver.find_element_by_name("jumin_tmp_05").send_keys(customer_num[2][0])
            driver.find_element_by_name("jumin_tmp_06").send_keys(customer_num[2][1])
            driver.find_element_by_name("jumin_tmp_07").send_keys(customer_num[2][2])

        driver.find_element_by_link_text("확인").click()
        driver.switch_to.window(driver.window_handles[0])

        printBillToPdf(driver, "{0:02d}. {1}.pdf".format(idx + 1, customer_num[0]))

    driver.switch_to.window(driver.window_handles[0])
    driver.quit()
