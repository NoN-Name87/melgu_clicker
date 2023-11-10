import person_parser as PP
import sys
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait   
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
from selenium.webdriver.firefox.options import Options
from openpyxl import Workbook
from openpyxl import load_workbook
from time import sleep
import os

def fill_user(row, browser):
    sleep(1.5)
    elem = WebDriverWait(browser, 300).until(lambda x: x.find_element(By.XPATH, '/html/body/div[2]/div/div[1]/div[2]/div/div/div[2]/div/div/div[1]/button'))
    send_xpath = '/html/body/div[2]/div/div[2]/div[2]/div/div[2]/div/div[11]/div/button'
    elem.click()
    sleep(1)
    last_name = browser.find_element(By.ID, "last_name")            
    last_name.send_keys(row["Фамилия"])
    sleep(1)
    first_name = browser.find_element(By.ID, "first_name")
    first_name.send_keys(row["Имя"])
    sleep(1)
    middle_name = browser.find_element(By.ID, "middle_name")
    middle_name.send_keys(row["Отчество"])
    sleep(1)
    email = browser.find_element(By.ID, "email")
    email.send_keys(row["Почта"])
    sleep(1)
    send_btn = browser.find_element(By.XPATH, send_xpath)
    send_btn.click()
    print('CHECK')
    try:
        print('CHECK 1')
        check_btn = browser.find_element(By.CSS_SELECTOR, '.pb-4 > button:nth-child(1)')
        print('SWAP')
        sleep(1)
        check_btn.click()
        sleep(1)
        first_name = browser.find_element(By.CSS_SELECTOR, '#first_name')
        first_name.clear()
        sleep(1)
        first_name.send_keys(row["Фамилия"])
        last_name = browser.find_element(By.CSS_SELECTOR, '#last_name')
        last_name.clear()
        sleep(1)
        last_name.send_keys(row["Имя"])
        sleep(1)
        send_btn.click()
    except Exception as ex:
        print('No Button')
        
    
def fill_browser_fields(result_dict, filename):
    try:
        options=Options()
        options.add_argument("-profile")
        options.add_argument("/home/vlados/snap/firefox/common/.mozilla/firefox/s9zkgp9s.default")
        browser = webdriver.Firefox(options=options)
        browser.get('https://melgu.diguniverse.ru/approval-personal-data')
        sleep(2)
        send_xpath = str()
        row = dict()
        # if(result_dict[0]["Статус"] == "Сотрудник"):
        #     elem = browser.find_element(By.XPATH, '/html/body/div[2]/div/div[1]/div[2]/div/div/div[2]/div/div/div[2]/button')
        #     send_xpath = '/html/body/div[2]/div/div[2]/div[2]/div/div[2]/div/div[10]/div/button'
        # else:
        #     elem = browser.find_element(By.XPATH, '/html/body/div[2]/div/div[1]/div[2]/div/div/div[2]/div/div/div[1]/button')
        #     send_xpath = '/html/body/div[2]/div/div[2]/div[2]/div/div[2]/div/div[11]/div/button'
        for row in result_dict:    
            row["Фамилия"].replace(' ', '')
            row["Имя"].replace(' ', '')
            row["Отчество"].replace(' ', '')
            row["Почта"].replace(' ', '')
            try:
                fill_user(row, browser)
            except Exception as ex:
                print("TIMEOUT 5 sec")
                sleep(10)
                browser.refresh()
                sleep(10)
                fill_user(row, browser)
            sleep(50)
            browser.refresh()
            PP.add_row(row)
            PP.delete_first_row(filename)
    except Exception as ex:
        print(ex)
    finally:
        exc_file = open("exception.txt", "w+")
        exc_file.write(f'ERROR, last ID is {row["ID"]}, {row["Имя"]}, {row["Фамилия"]}, {row["Отчество"]}')
        os.rename('test/dump.xlsx', f'test/dump_{int(row["ID"])}.xlsx')
        exc_file.close()
        browser.close()
        browser.quit()

if __name__ == "__main__":
    filename = sys.argv[1]
    path = os.path.join('test', filename)
    result_dict = PP.parse_person(path)
    person_headers = ['ID', 'Имя', 'Фамилия', 'Отчество', 'Почта']
    wb = Workbook()
    ws = wb.active
    ws.append(person_headers)
    wb.save(os.path.join('test', 'dump.xlsx'))
    fill_browser_fields(result_dict, path)