from format_selenium_input_data import normalize_excel_sheet
from format_selenium_input_data import get_contact_list
from format_selenium_input_data import find_labels
import time
import random
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support import ui
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import sys
import datetime
import copy

global first_contact


def page_is_loaded(driver):
    return driver.find_element_by_tag_name("body") != None

def close_blank_page(driver, wait, link):
    # open webpage
    driver.get(link)
    wait.until(page_is_loaded)

    # close the blank page that opens default with selenium and assign a main window 
    windows = driver.window_handles
    for window in windows:
        driver.switch_to.window(window)
        if len(driver.find_elements_by_css_selector("*")) >= 10:
            main_window = window
        else:
            driver.switch_to.window(window)
            driver.close()
    driver.switch_to.window(main_window)

    return main_window

def login_to_missionhub(driver, wait, main):
    time.sleep(1)
    
    # sign into facebook
    driver.find_element_by_xpath('/html/body/ui-view/app/section/ui-view/sign-in/div/div[3]/a[2]').click()
    wait.until(page_is_loaded)
    # switch to newly opened window
    windows = driver.window_handles
    driver.switch_to.window(windows[-1])
    # type login info into fb
    driver.find_element_by_xpath('//*[@id="email"]').send_keys('connorivy15@gmail.com')
    driver.find_element_by_xpath('//*[@id="pass"]').send_keys('September15!')
    driver.find_element_by_xpath('//*[@id="u_0_0"]').click()

    driver.switch_to.window(main)
    wait.until(page_is_loaded)
    
    time.sleep(1)

def add_new_contact(driver, wait, contact_info, user_labels):
    global first_contact
    
    user_labels_copy = copy.copy(user_labels)
    time.sleep(2)

    # add person btn
    try:
        driver.find_element_by_xpath('/html/body/ui-view/app/section/ui-view/my-people-dashboard/div/div[1]/organization/accordion/div[1]/accordion-header/div/div[2]/ng-md-icon[1]').click()
    except:
        time.sleep(4)
        driver.find_element_by_xpath('/html/body/ui-view/app/section/ui-view/my-people-dashboard/div/div[1]/organization/accordion/div[1]/accordion-header/div/div[2]/ng-md-icon[1]').click()
    wait.until(page_is_loaded)

    if contact_info[0] != None:
        driver.find_element_by_xpath('/html/body/div[1]/div/div/person-page/async-content/div/div[1]/person-profile/form/div[6]/div[1]/label/input').send_keys(contact_info[0])
    if contact_info[1] != None:
        driver.find_element_by_xpath('/html/body/div[1]/div/div/person-page/async-content/div/div[1]/person-profile/form/div[6]/div[2]/label/input').send_keys(contact_info[1])
    if contact_info[2] != None:
        driver.find_element_by_xpath('/html/body/div[1]/div/div/person-page/async-content/div/div[1]/person-profile/form/div[6]/div[5]/div/label/div[2]/input').send_keys(contact_info[2])

    # unassign the contact to yourself by default
    try:
        driver.find_element_by_xpath('/html/body/div[1]/div/div/person-page/async-content/div/div[1]/person-profile/form/div[3]/label/assigned-people-select/div/div[1]/span/span/span/span[1]').click()
    except:
        print("Expect to find someone preassigned to", contact_info[0], ' ', contact_info[1], " but didn't")

    # male 
    if contact_info[3] == 'male':
        try:
            driver.find_element_by_xpath('/html/body/div[1]/div/div/person-page/async-content/div/div[1]/person-profile/form/div[6]/div[3]/label[1]/input').click()
        except:
            print('Trouble clicking the gender option for ', contact_info[0], ' ', contact_info[1])
        time.sleep(2)
    
    # female
    elif contact_info[3] == 'female':
        try:
            driver.find_element_by_xpath('/html/body/div[1]/div/div/person-page/async-content/div/div[1]/person-profile/form/div[6]/div[3]/label[2]/input').click()
        except:
            print('Trouble clicking the gender option for ', contact_info[0], ' ', contact_info[1])
    
    # other
    elif contact_info[3] == 'other':
        try:
            driver.find_element_by_xpath('/html/body/div[1]/div/div/person-page/async-content/div/div[1]/person-profile/form/div[6]/div[3]/label[3]/input').click()
        except:
            print('Trouble clicking the gender option for ', contact_info[0], ' ', contact_info[1])

    # add label button
    driver.find_element_by_xpath('/html/body/div[1]/div/div/person-page/async-content/div/div[1]/person-profile/form/div[1]/div[1]/div[1]/ng-md-icon').click()
    availible_labels = driver.find_element_by_xpath('//*[@id="modal-body"]/multiselect-list/ul')

    a = datetime.datetime.now()
    list_elements = availible_labels.find_elements_by_xpath('.//*')
    for child in range (0,len(list_elements),3):  
        if list_elements[child].text.lower() in user_labels:
            user_labels_copy.remove(list_elements[child].text.lower())
            list_elements[child].find_element_by_css_selector('span[class=ng-binding]').click()

    b = datetime.datetime.now()
    print('time', b-a)

    # if this is the first contact, then check if the label was added
    # if not, the create a new label and then call the function again with the same contact info
    if first_contact:
        first_contact = False
        if user_labels_copy != []:
            add_labels_to_mh(driver, wait, user_labels_copy)
            add_new_contact(driver, wait, contact_info, user_labels)
        else:
            # the OK btn
            driver.find_element_by_xpath('/html/body/div[1]/div/div/edit-group-or-label-assignments/div[3]/button[2]/span').click()

            # save btn
            driver.find_element_by_xpath('/html/body/div[1]/div/div/person-page/async-content/div/div[2]/button').click()
            wait.until(page_is_loaded)
    else:
        # the OK btn
        driver.find_element_by_xpath('/html/body/div[1]/div/div/edit-group-or-label-assignments/div[3]/button[2]/span').click()

        # save btn
        driver.find_element_by_xpath('/html/body/div[1]/div/div/person-page/async-content/div/div[2]/button').click()
        wait.until(page_is_loaded)

def add_labels_to_mh(driver, wait, user_labels):
    # the OK btn
    driver.find_element_by_xpath('/html/body/div[1]/div/div/edit-group-or-label-assignments/div[3]/button[2]/span').click()

    # x at the top right
    driver.find_element_by_xpath('/html/body/div[1]/div/div/person-page/async-content/div/header/div[2]/div[1]/a').click()

    # ok btn on the are you sure page
    driver.find_element_by_xpath('/html/body/div[1]/div/div/div/div[3]/button[2]').click()

    # click on 'cru @ the university of texas'
    driver.find_element_by_xpath('/html/body/ui-view/app/section/ui-view/my-people-dashboard/div/div[1]/organization/accordion/div[1]/accordion-header/div/div[1]/h2').click()
    wait.until(page_is_loaded)

    # hover over the tools dropdown menu
    menu = driver.find_element_by_xpath('/html/body/ui-view/app/section/ui-view/my-organizations-dashboard/div/ui-view/organization-overview/async-content/div/div/div[2]/div[7]/div')
    ActionChains(driver).move_to_element(menu).perform()

    # click on 'manage labels'
    driver.find_element_by_xpath('/html/body/ui-view/app/section/ui-view/my-organizations-dashboard/div/ui-view/organization-overview/async-content/div/div/div[2]/div[7]/div/ul/li[3]/a').click()
    
    # click the plus btn to add label
    driver.find_element_by_xpath('/html/body/ui-view/app/section/ui-view/my-organizations-dashboard/div/ui-view/organization-overview/async-content/div/div/div[3]/ui-view/organization-overview-labels/div[1]/div[2]/icon-button/ng-md-icon').click()

    # type new label in box for each element left in user labels
    for x in user_labels:
        driver.find_element_by_xpath('//*[@id="modal-body"]/div/label/input').send_keys(x)

    # click the okay label
    driver.find_element_by_xpath('/html/body/div[1]/div/div/edit-label/div[3]/button[2]').click()

    # go back to the people tab
    driver.get('https://campuscontacts.cru.org/people')
    wait.until(page_is_loaded)

def main():
    global first_contact
    
    chromedriver = "chromedriver.exe"
    driver = webdriver.Chrome(chromedriver)
    driver.implicitly_wait(10)
    wait = ui.WebDriverWait(driver, 10)
    link = 'https://campuscontacts.cru.org/sign-in'

    normalize_excel_sheet()
    contact_list = get_contact_list()
    labels = find_labels()
    main_window = close_blank_page(driver, wait, link)
    login_to_missionhub(driver, wait, main_window)

    first_contact = True

    for contact in contact_list:
        add_new_contact(driver, wait, contact, labels)

    print('all done :)')
    time.sleep(60)


if __name__ == "__main__":
    main()