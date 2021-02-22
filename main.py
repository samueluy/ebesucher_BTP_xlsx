import openpyxl
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from datetime import datetime


def get_current_btp(username, password):
    PATH = "C:\Program Files (x86)\chromedriver.exe" # Change to path of web driver
    driver = webdriver.Chrome(PATH) # Change according to web driver

    driver.get("https://www.ebesucher.com/statistik")

    username_box = driver.find_element_by_id("LoginForm_login_name")
    username_box.send_keys(username)
    password_box = driver.find_element_by_id("LoginForm_login_password")
    password_box.send_keys(password)
    password_box.send_keys(Keys.RETURN)

    total = driver.find_elements_by_xpath("""//*[@id="statstable"]/tbody/tr[1]/td[3]""")
    current_btp = total[0].text.replace(" BTP", "").replace(",", "")[:-3]

    driver.quit()
    return current_btp


def main():
    current_btp = get_current_btp(username="USERNAME" , password="PASSWORD") #Input username and password
    now = datetime.now().strftime("%d/%m/%Y %H:%M")

    excel_file = "ebesucher.xlsx"

    wb = openpyxl.load_workbook(excel_file)
    sheet = wb["Earnings"]

    for cell in sheet['A2':'A40']:
        if cell[0].value == None:
            cell[0].value = now
            break

    for cell in sheet['C2':'C40']:
        if cell[0].value == None:
            cell[0].value = float(current_btp)
            wb.save(excel_file)
            print(now, current_btp)

            quit()


main()
