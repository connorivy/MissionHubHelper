import xlrd
import time
import random
from selenium import webdriver
from selenium.webdriver.support import ui
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import gspread
from oauth2client.service_account import ServiceAccountCredentials


def page_is_loaded(driver):
    return driver.find_element_by_tag_name("body") != None

def modal_page_is_loaded(driver):
    return driver.find_element_by_id("modal-body") != None

def get_user_info():
    user_info = []
    filepath="user_info.txt"
    with open(filepath, encoding="utf8") as fp:  
        for line in fp:
            user_info.extend(line.strip().split(': '))

    print(user_info)
    return user_info

def get_contacts_excel():
    # Reading an excel file using Python 
    import xlrd 

    # Give the location of the file 
    loc = ("./contacts.xlsx") 

    # To open Workbook 
    wb = xlrd.open_workbook(loc) 
    sheet = wb.sheet_by_index(0) 

    small_list = []
    big_list = []
    
    print(sheet.nrows)

    for row in range(sheet.nrows):
        for col in range(3):
            small_list.append(sheet.cell_value(row, col))
        big_list.append(small_list)
        small_list = []
        
    return big_list

def login_to_missionhub(driver, wait, main):
    time.sleep(1)

    driver.find_element_by_xpath('//*[@id="menu-item-1971"]/a').click()
    wait.until(page_is_loaded)

    driver.find_element_by_xpath('/html/body/ui-view/app/section/ui-view/sign-in/div/p[2]/a').click()
    wait.until(page_is_loaded)

    windows = driver.window_handles
    driver.switch_to.window(windows[-1])

    driver.find_element_by_xpath('//*[@id="email"]').send_keys('connorivy15@gmail.com')
    driver.find_element_by_xpath('//*[@id="pass"]').send_keys('September15!')
    driver.find_element_by_xpath('//*[@id="u_0_0"]').click()

    driver.switch_to.window(main)
    wait.until(page_is_loaded)
    time.sleep(5)

def get_contacts_google(info):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json',scope)
    client = gspread.authorize(creds)

    sh = client.open(info[5])

def add_new_contact(driver, wait, small_list):
    driver.find_element_by_xpath('/html/body/ui-view/app/section/ui-view/my-people-dashboard/div/div[1]/organization/accordion/div[1]/accordion-header/div/div[2]/ng-md-icon[1]').click()
    wait.until(page_is_loaded)

    driver.find_element_by_xpath('/html/body/div[1]/div/div/person-page/async-content/div/div[1]/person-profile/form/div[6]/div[1]/label/input').send_keys(small_list[0])
    driver.find_element_by_xpath('/html/body/div[1]/div/div/person-page/async-content/div/div[1]/person-profile/form/div[6]/div[2]/label/input').send_keys(small_list[1])
    driver.find_element_by_xpath('/html/body/div[1]/div/div/person-page/async-content/div/div[1]/person-profile/form/div[6]/div[5]/div/label/div[2]/input').send_keys(small_list[2])

    driver.find_element_by_xpath('/html/body/div[1]/div/div/person-page/async-content/div/div[1]/person-profile/form/div[3]/label/assigned-people-select/div/div[1]/span/span/span/span[1]').click()

    driver.find_element_by_xpath('/html/body/div[1]/div/div/person-page/async-content/div/div[1]/person-profile/form/div[1]/div[1]/div[1]/ng-md-icon').click()
    time.sleep(.75)
    driver.find_element_by_xpath('//*[@id="modal-body"]/multiselect-list/ul/li[85]/span/span').click()
    driver.find_element_by_xpath('/html/body/div[1]/div/div/edit-group-or-label-assignments/div[3]/button[2]/span').click()
    time.sleep(.75)

    driver.find_element_by_xpath('/html/body/div[1]/div/div/person-page/async-content/div/div[2]/button').click()
    wait.until(page_is_loaded)

def sort_contacts(driver, wait, all_contacts):
    failed_contacts = []
    filepath="failed_contacts.txt"
    with open(filepath, encoding="utf8") as fp:  
        for line in fp:
            failed_contacts.extend(line.strip().split(', '))

    for contact in all_contacts:
        try:
            if (len(contact[2]) > 12):
                print('number wrong!!!')
                for num in range(len(failed_contacts)//3):
                    if (failed_contacts[(num+1)*3-1] == contact[2]):
                        continue
                    else:
                        with open("failed_contacts.txt", "a") as text_file:
                            text_file.write("%s, %s, %s\n" % (str(contact[0]), str(contact[1]), str(contact[2])))
        except:
            print()

        else:
            try:
                add_new_contact(driver, wait, contact)
                time.sleep(1)
            except:
                time.sleep(2)

def main():
    chromedriver = "chromedriver.exe"
    driver = webdriver.Chrome(chromedriver)
    wait = ui.WebDriverWait(driver, 10)
    link = 'https://get.missionhub.com/'

    driver.get(link)
    wait.until(page_is_loaded)

    windows = driver.window_handles
    print(windows)
    for window in windows:
        driver.switch_to.window(window)
        if len(driver.find_elements_by_css_selector("*")) >= 10:
            main_window = window
        else:
            driver.switch_to.window(window)
            driver.close()
    driver.switch_to.window(main_window)
    
    # info = get_user_info()
    # import_contacts(info)
    all_contacts = get_contacts_excel()
    login_to_missionhub(driver, wait, main_window)
    sort_contacts(driver, wait, all_contacts)
    time.sleep(60)


if __name__ == "__main__":
    main()