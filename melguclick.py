import person_parser as PP
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
from selenium.webdriver.firefox.options import Options
from time import sleep

def fill_browser_fields(result_dict):
    options=Options()
    options.add_argument("-profile")
    options.add_argument("/home/vlados/snap/firefox/common/.mozilla/firefox/s9zkgp9s.default")
    browser = webdriver.Firefox(options=options)
    browser.get('https://melgu.diguniverse.ru/approval-personal-data')
    sleep(2)
    if(result_dict[0]["Статус"] == "Сотрудник"):
        elem = browser.find_element(By.XPATH, '//button[@wire:click="initEmployeeForm()"]')
    else:
        elem = browser.find_element(By.XPATH, '//button[@wire:click="initStudentForm()"]')
    elem.click()
    sleep(2)

if __name__ == "__main__":
    result_dict = PP.parse_person('test/persons.xlsx')
    print(result_dict)
