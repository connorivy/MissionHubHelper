from format_selenium_input_data import normalize_excel_sheet
from format_selenium_input_data import get_contact_list
from format_selenium_input_data import find_labels
import time
import random
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
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
import base64
from getpass import getpass

first_contact = True

def is_headless():
    headless = 'change me'
    print('\n\n\n')
    while headless.lower() != 'y' and headless.lower() != 'n':
        headless = input('Would you like to run a headless browser? (y/n)     ')
    
    if headless == 'y':
        return True
    else:
        return False

def retrieve_login_info():
    login_info = read_in_login_info()

    if login_info[1] == '' or login_info[3] == '' or login_info[5] == '':
        reset_login_info()
        login_info = read_in_login_info()
    
    login_info[1] = base64.b64decode(login_info[1].encode("utf-8")).decode("utf-8")
    login_info[3] = base64.b64decode(login_info[3].encode("utf-8")).decode("utf-8")
    login_info[5] = base64.b64decode(login_info[5].encode("utf-8")).decode("utf-8")

    return login_info

def read_in_login_info():
    with open('./supporting_files/login_info.txt') as f:
        temp_info = f.readlines()
    f.close()

    # split based on the ':' character
    temp_info = [x.split(':') for x in temp_info]

    login_info = []
    # remove whitespace and /n
    for x in temp_info:
        for y in x:
            login_info.append(y.strip())

    return login_info

def reset_login_info():
    login_info = ['email or facebook', '', 'username', '', 'password', '']

    print('\n\n\n\n')
    while login_info[1].lower() != 'f' and login_info[1].lower() != 'e':
        login_info[1] = input('\nLogin via email or Facebook? [E/F]     ')
    
    login_info[3] = input('Please input your username:     ')
    login_info[5] = getpass('Please input your password:     ')

    file_info = ['','','']
    file_info[0] = "email or facebook:" + base64.b64encode(login_info[1].encode("utf-8")).decode("utf-8") + "\n"
    file_info[1] = "username:" + base64.b64encode(login_info[3].encode("utf-8")).decode("utf-8") + "\n"
    file_info[2] = "password:" + base64.b64encode(login_info[5].encode("utf-8")).decode("utf-8")
    with open('./supporting_files/login_info.txt', 'w') as f:
        f.writelines(file_info)
    f.close()

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

    # returns login info in the order ['text', 'f or e', 'text', 'username', 'text', 'password']
    login_info = retrieve_login_info()

    if login_info[1].lower() == 'f':
        # sign into facebook btn 
        try_to_click(driver, '/html/body/ui-view/app/section/ui-view/sign-in/div/div[3]/a[2]')
        wait.until(page_is_loaded)
        # switch to newly opened window
        windows = driver.window_handles
        driver.switch_to.window(windows[-1])
        # type login info into fb
        try_to_send_keys(driver, '//*[@id="email"]', login_info[3])
        try_to_send_keys(driver, '//*[@id="pass"]', login_info[5])
        try_to_click(driver, '//*[@id="u_0_0"]')
    
    else:
        # sign in with email btn
        try_to_click(driver, '/html/body/ui-view/app/section/ui-view/sign-in/div/div[3]/a[1]')
        wait.until(page_is_loaded)

        # type login info
        try_to_send_keys(driver, '//*[@id="username"]', login_info[3])
        try_to_send_keys(driver, '//*[@id="password"]', login_info[5])
        try_to_click(driver, '//*[@id="login_form"]/div[3]/button')

    driver.switch_to.window(main)
    wait.until(page_is_loaded)
    
    time.sleep(1)

    # Cru at university of Texas
    try_to_click(driver, '/html/body/ui-view/app/section/ui-view/my-people-dashboard/div/div[1]/organization/accordion/div[1]/accordion-header/div/div[1]/h2')

    # contacts 
    try_to_click(driver, '/html/body/ui-view/app/section/ui-view/my-organizations-dashboard/div/ui-view/organization-overview/async-content/div/div/div[2]/div[4]/span')

def add_new_contact(driver, wait, contact_info, user_labels):
    global first_contact
    
    # add the year to the contact's labels
    if contact_info[5] != None:
        user_labels.append(contact_info[5])

    # add person btn
    try_to_click(driver, '/html/body/ui-view/app/section/ui-view/my-organizations-dashboard/div/ui-view/organization-overview/async-content/div/div/div[3]/ui-view/organization-overview-people/people-screen/div/div[2]/div/div[2]/div[1]/div/div[1]/div[2]/icon-button')
    wait.until(page_is_loaded)

    fill_in_contact(driver, wait, contact_info, user_labels)

    assign_gender(driver, wait, contact_info)

    textbox = try_to_find_element(driver, '/html/body/ui-view/app/section/ui-view/my-organizations-dashboard/div/ui-view/organization-overview/async-content/div/div/div[3]/ui-view/organization-overview-people/people-screen/div/div[2]/div/div[1]/people-filters-panel/div/div[1]/input')
    textbox.clear()
    
def fill_in_contact(driver, wait, contact_info, user_labels):
    global first_contact
    user_labels_copy = copy.copy(user_labels)

    # fill in first, last, and phone
    if contact_info[1] != None:
        try_to_send_keys(driver, '/html/body/div[1]/div/div/person-page/async-content/div/div[1]/person-profile/form/div[6]/div[1]/label/input', contact_info[1])
    if contact_info[2] != None:
        try_to_send_keys(driver, '/html/body/div[1]/div/div/person-page/async-content/div/div[1]/person-profile/form/div[6]/div[2]/label/input', contact_info[2])
    if contact_info[3] != None:
        try_to_send_keys(driver, '/html/body/div[1]/div/div/person-page/async-content/div/div[1]/person-profile/form/div[6]/div[5]/div/label/div[2]/input', contact_info[3])

    # click out of the name field
    try_to_click(driver, '/html/body/div[1]/div/div/person-page/async-content/div/div[1]/person-profile/form/div[6]')

    # male 
    if contact_info[4] == 'male':
        try_to_click(driver, '/html/body/div[1]/div/div/person-page/async-content/div/div[1]/person-profile/form/div[6]/div[3]/label[1]/input')
    
    # female
    elif contact_info[4] == 'female':
        try_to_click(driver, '/html/body/div[1]/div/div/person-page/async-content/div/div[1]/person-profile/form/div[6]/div[3]/label[2]/input')
    
    # other
    elif contact_info[4] == 'other':
        try_to_click(driver, '/html/body/div[1]/div/div/person-page/async-content/div/div[1]/person-profile/form/div[6]/div[3]/label[3]/input')

    # add label button
    try_to_click(driver, '/html/body/div[1]/div/div/person-page/async-content/div/div[1]/person-profile/form/div[1]/div[1]/div[1]/ng-md-icon')
    availible_labels = driver.find_element_by_xpath('//*[@id="modal-body"]/multiselect-list/ul')

    # parse through list of current labels, add label if it exists
    list_elements = availible_labels.find_elements_by_xpath('.//*')
    for child in range (0,len(list_elements),3):  
        if list_elements[child].text.lower() in user_labels:
            user_labels_copy.remove(list_elements[child].text.lower())
            list_elements[child].find_element_by_css_selector('span[class=ng-binding]').click()

    # if this is the first contact, then check if the label was added
    # if it wasn't added then create a new label and then call the function again with the same contact info
    if first_contact:
        first_contact = False
        if user_labels_copy != []:
            add_labels_to_mh(driver, wait, user_labels_copy)
            add_new_contact(driver, wait, contact_info, user_labels)
        else:
            # the OK btn
            try_to_click(driver, '/html/body/div[1]/div/div/edit-group-or-label-assignments/div[3]/button[2]/span')

            # save btn
            try_to_click(driver, '/html/body/div[1]/div/div/person-page/async-content/div/div[2]/button')
            wait.until(page_is_loaded)
            print(contact_info[1], contact_info[2], 'added successfully')
    else:
        # the OK btn
        try_to_click(driver, '/html/body/div[1]/div/div/edit-group-or-label-assignments/div[3]/button[2]/span')

        # save btn
        try_to_click(driver, '/html/body/div[1]/div/div/person-page/async-content/div/div[2]/button')
        wait.until(page_is_loaded)
        print(contact_info[1], contact_info[2], 'added successfully')

def assign_gender(driver, wait, contact_info):
    stop = False

    # # search for the person who was just added
    # if contact_info[2] != None:
    #     try_to_send_keys(driver, '/html/body/ui-view/app/section/ui-view/my-organizations-dashboard/div/ui-view/organization-overview/async-content/div/div/div[3]/ui-view/organization-overview-people/people-screen/div/div[2]/div/div[1]/people-filters-panel/div/div[1]/input', contact_info[1] + " " + contact_info[2])
    # else:
    #     try_to_send_keys(driver, '/html/body/ui-view/app/section/ui-view/my-organizations-dashboard/div/ui-view/organization-overview/async-content/div/div/div[3]/ui-view/organization-overview-people/people-screen/div/div[2]/div/div[1]/people-filters-panel/div/div[1]/input', contact_info[1])
    try_to_send_keys(driver, '/html/body/ui-view/app/section/ui-view/my-organizations-dashboard/div/ui-view/organization-overview/async-content/div/div/div[3]/ui-view/organization-overview-people/people-screen/div/div[2]/div/div[1]/people-filters-panel/div/div[1]/input', contact_info[3])
    time.sleep(1.5)

    # try to find the first label that pops up, if no labels popped up
    try:
        gender = try_to_find_element(driver, '/html/body/ui-view/app/section/ui-view/my-organizations-dashboard/div/ui-view/organization-overview/async-content/div/div/div[3]/ui-view/organization-overview-people/people-screen/div/div[2]/div/div[2]/div[2]/div/ministry-view-person/div/div[2]/span')
    except:
        if contact_info[2] != None:
            print('\n\n\n***** PROBLEM ADDING ' + contact_info[1] + ' ' + contact_info[2] + ' *****')
        else:
            print('*****PROBLEM ADDING ' + contact_info[1] +  ' *****')

        # dont keep going
        stop = True

    if not stop:
        if '-' in gender.text:
            try_to_click(driver, '/html/body/ui-view/app/section/ui-view/my-organizations-dashboard/div/ui-view/organization-overview/async-content/div/div/div[3]/ui-view/organization-overview-people/people-screen/div/div[2]/div/div[2]/div[2]/div/ministry-view-person[1]/div/div[1]/div[2]/a')

            # male 
            if contact_info[4] == 'male':
                try_to_click(driver, '/html/body/ui-view/app/section/ui-view/my-organizations-dashboard/div/ui-view/organization-overview/async-content/div/div/div[3]/ui-view/organization-overview-people/people-screen/div/div[2]/div/ui-view/person-page/async-content/div/div/person-profile/form/div[6]/div[3]/label[1]/input')
            
            # female
            elif contact_info[4] == 'female':
                try_to_click(driver, '/html/body/ui-view/app/section/ui-view/my-organizations-dashboard/div/ui-view/organization-overview/async-content/div/div/div[3]/ui-view/organization-overview-people/people-screen/div/div[2]/div/ui-view/person-page/async-content/div/div/person-profile/form/div[6]/div[3]/label[2]/input')
            
            # other
            elif contact_info[4] == 'other':
                try_to_click(driver, '/html/body/ui-view/app/section/ui-view/my-organizations-dashboard/div/ui-view/organization-overview/async-content/div/div/div[3]/ui-view/organization-overview-people/people-screen/div/div[2]/div/ui-view/person-page/async-content/div/div/person-profile/form/div[6]/div[3]/label[3]/input')

            # click the x
            try_to_click(driver, '/html/body/ui-view/app/section/ui-view/my-organizations-dashboard/div/ui-view/organization-overview/async-content/div/div/div[3]/ui-view/organization-overview-people/people-screen/div/div[2]/div/ui-view/person-page/async-content/div/header/div[2]/div[1]/a')             
        
        if contact_info[0] != None:
            # click the person
            try_to_click(driver, '/html/body/ui-view/app/section/ui-view/my-organizations-dashboard/div/ui-view/organization-overview/async-content/div/div/div[3]/ui-view/organization-overview-people/people-screen/div/div[2]/div/div[2]/div[2]/div/ministry-view-person[1]/div/div[1]/div[2]/a')

            # history
            try_to_click(driver, '/html/body/ui-view/app/section/ui-view/my-organizations-dashboard/div/ui-view/organization-overview/async-content/div/div/div[3]/ui-view/organization-overview-people/people-screen/div/div[2]/div/ui-view/person-page/async-content/div/header/div[2]/div[3]/div[2]')

            # notes
            try_to_click(driver, '/html/body/ui-view/app/section/ui-view/my-organizations-dashboard/div/ui-view/organization-overview/async-content/div/div/div[3]/ui-view/organization-overview-people/people-screen/div/div[2]/div/ui-view/person-page/async-content/div/div/person-history/async-content/div/div[1]/span[3]')

            # add note
            try_to_click(driver, '/html/body/ui-view/app/section/ui-view/my-organizations-dashboard/div/ui-view/organization-overview/async-content/div/div/div[3]/ui-view/organization-overview-people/people-screen/div/div[2]/div/ui-view/person-page/async-content/div/div/person-history/async-content/div/div[2]/div[3]/ng-md-icon')

            #comment only
            try_to_click(driver, '/html/body/ui-view/app/section/ui-view/my-organizations-dashboard/div/ui-view/organization-overview/async-content/div/div/div[3]/ui-view/organization-overview-people/people-screen/div/div[2]/div/ui-view/person-page/async-content/div/div/person-history/async-content/div/div[2]/div[2]/div/div[1]')

            # write comment
            try_to_send_keys(driver, '/html/body/ui-view/app/section/ui-view/my-organizations-dashboard/div/ui-view/organization-overview/async-content/div/div/div[3]/ui-view/organization-overview-people/people-screen/div/div[2]/div/ui-view/person-page/async-content/div/div/person-history/async-content/div/div[2]/form/div/textarea', contact_info[0])

            # check mark
            try_to_click(driver, '/html/body/ui-view/app/section/ui-view/my-organizations-dashboard/div/ui-view/organization-overview/async-content/div/div/div[3]/ui-view/organization-overview-people/people-screen/div/div[2]/div/ui-view/person-page/async-content/div/div/person-history/async-content/div/div[2]/form/button/ng-md-icon')

            # click the x
            try_to_click(driver, '/html/body/ui-view/app/section/ui-view/my-organizations-dashboard/div/ui-view/organization-overview/async-content/div/div/div[3]/ui-view/organization-overview-people/people-screen/div/div[2]/div/ui-view/person-page/async-content/div/header/div[2]/div[1]/a')             
        

def add_labels_to_mh(driver, wait, user_labels):
    # the OK btn
    try_to_click(driver, '/html/body/div[1]/div/div/edit-group-or-label-assignments/div[3]/button[2]')

    # x at the top right
    try_to_click(driver, '/html/body/div[1]/div/div/person-page/async-content/div/header/div[2]/div[1]/a')

    # ok btn on the are you sure page
    try_to_click(driver, '/html/body/div[1]/div/div/div/div[3]/button[2]')

    # hover over the tools dropdown menu
    menu = try_to_find_element(driver, '/html/body/ui-view/app/section/ui-view/my-organizations-dashboard/div/ui-view/organization-overview/async-content/div/div/div[2]/div[7]/div')
    ActionChains(driver).move_to_element(menu).perform()

    # click on 'manage labels'
    try_to_click(driver, '/html/body/ui-view/app/section/ui-view/my-organizations-dashboard/div/ui-view/organization-overview/async-content/div/div/div[2]/div[7]/div/ul/li[3]/a')
    
    
    for x in user_labels:
        # click the plus btn to add label
        try_to_click(driver, '/html/body/ui-view/app/section/ui-view/my-organizations-dashboard/div/ui-view/organization-overview/async-content/div/div/div[3]/ui-view/organization-overview-labels/div[1]/div[2]/icon-button/ng-md-icon')

        # type new label in box for each element left in user labels
        try_to_send_keys(driver, '//*[@id="modal-body"]/div/label/input', x)

        # click the okay label
        try_to_click(driver, '/html/body/div[1]/div/div/edit-label/div[3]/button[2]')


    # go back to the contacts tab
    try_to_click(driver, '/html/body/ui-view/app/section/ui-view/my-organizations-dashboard/div/ui-view/organization-overview/async-content/div/div/div[2]/div[4]')

def try_to_click(driver, xpath):
    try:
        driver.find_element_by_xpath(xpath).click()
    except:
        time.sleep(3)
        driver.find_element_by_xpath(xpath).click()

def try_to_send_keys(driver, xpath, keys):
    try:
        driver.find_element_by_xpath(xpath).send_keys(keys)
    except:
        time.sleep(3)
        driver.find_element_by_xpath(xpath).send_keys(keys)

def try_to_find_element(driver, xpath):
    try:
        element = driver.find_element_by_xpath(xpath)
    except:
        time.sleep(3)
        element = driver.find_element_by_xpath(xpath)

    return element


def main():
    global first_contact

    headless = is_headless()
    chrome_options = Options()
    chrome_options.headless = headless
    
    chromedriver = "supporting_files/chromedriver.exe"
    # driver = webdriver.Chrome(chromedriver)
    driver = webdriver.Chrome(ChromeDriverManager().install())
    driver.implicitly_wait(10)
    wait = ui.WebDriverWait(driver, 5)
    link = 'https://campuscontacts.cru.org/sign-in'

    normalize_excel_sheet()
    # contact list in the form [notes, first, last, phone, gender, year]
    contact_list = get_contact_list()
    print('contact list', contact_list)
    labels = find_labels()
    main_window = close_blank_page(driver, wait, link)

    start = datetime.datetime.now()
    login_to_missionhub(driver, wait, main_window)

    for contact in contact_list:
        add_new_contact(driver, wait, contact, labels)

    finish = datetime.datetime.now()
    print('all done :)')
    print('Added', len(contact_list), 'contacts in', finish - start)
    time.sleep(5)


if __name__ == "__main__":
    main()