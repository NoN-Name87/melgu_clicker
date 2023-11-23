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
    bad_str = "nan"
    while True:
        try:
            sleep(0.5)
            elem = WebDriverWait(browser, 10).until(lambda x: x.find_element(By.XPATH, '/html/body/div[2]/div/div[1]/div[2]/div/div/div[2]/div/div/div[1]/button'))
            send_xpath = '/html/body/div[2]/div/div[2]/div[2]/div/div[2]/div/div[11]/div/button'
            elem.click()
            sleep(0.5)
            last_name = browser.find_element(By.ID, "last_name")            
            last_name.send_keys(row["Фамилия"])
            sleep(0.5)
            first_name = browser.find_element(By.ID, "first_name")
            first_name.send_keys(row["Имя"])
            sleep(0.5)
            if type(row["Отчество"]) != float:
                middle_name = browser.find_element(By.ID, "middle_name")
                middle_name.send_keys(row["Отчество"])
            sleep(0.5)
            email = browser.find_element(By.ID, "email")
            email.send_keys(row["Почта"])
            sleep(0.5)
            send_btn = browser.find_element(By.XPATH, send_xpath)
            send_btn.click()
            print('CHECK')
            break
        except Exception as ex:
            print(ex)
            print('WAIT 1 min')
            sleep(60)
            browser.refresh()
    try:
        print('CHECK 1')
        check_btn = browser.find_element(By.CSS_SELECTOR, '.pb-4 > button:nth-child(1)')
        print('SWAP')
        sleep(0.5)
        check_btn.click()
        sleep(0.5)
        first_name = browser.find_element(By.CSS_SELECTOR, '#first_name')
        first_name.clear()
        sleep(0.5)
        first_name.send_keys(row["Фамилия"])
        last_name = browser.find_element(By.CSS_SELECTOR, '#last_name')
        last_name.clear()
        sleep(0.5)
        last_name.send_keys(row["Имя"])
        sleep(0.5)
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
        for row in result_dict:
            try:    
                row["Фамилия"].replace(" ", "")
                row["Имя"].replace(" ", "")
                row["Почта"].replace(" ", "")
                row["Отчество"].replace(" ", "")
            except Exception as ex:
                print("Empty block")
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