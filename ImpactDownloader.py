# Impact Downloader script by Nick Furlo 5/21/18
import configparser
import csv
import glob
import itertools
import os
import threading
import time
import tkinter
from datetime import datetime
from pathlib import Path
from tkinter import *
from tkinter import messagebox

from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from tqdm import tqdm
from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.common.keys import Keys


# Create or load config file and assign variables.
def load_config():
    # If there is no config file, make one.
    config = configparser.ConfigParser()
    my_file = Path("ImpactDownloader.config")
    if not my_file.is_file():
        file = open("ImpactDownloader.config", "w+")
        file.write(
            "[DEFAULT]\nUSERNAME = \nPASSWORD = \nINPUT_CSV_FILE_PATH = \nNUMBER_OF_ROWS = \nDOWNLOAD_LOCATION =")
        messagebox.showinfo("Warning", "Config file created. Please add a CSV file path and relaunch the program.")
        driver.close()
        sys.exit()

    # Read config and set variables
    config.read('ImpactDownloader.config')
    global username, password, input_file_path, number_of_rows, download_file_path
    username = config['DEFAULT']['USERNAME']
    password = config['DEFAULT']['PASSWORD']
    input_file_path = config['DEFAULT']['INPUT_CSV_FILE_PATH']
    number_of_rows = int(config['DEFAULT']['NUMBER_OF_ROWS'])
    download_file_path = config['DEFAULT']['DOWNLOAD_LOCATION']

    # End if there is no CSV path
    if input_file_path is "":
        messagebox.showinfo("Warning", "Please enter a CSV file path into the config file and relaunch the program.")
        driver.close()
        sys.exit()


# load CSV contents of first 2 columns row 1 through number_of_rows into a dictionary
def load_csv():
    with open(input_file_path, mode='r') as infile:
        urls = {rows[0]: rows[1] for rows in itertools.islice(csv.reader(infile), 1, number_of_rows)}
        return urls


# Downloads csv file for insights pages
def download_insight_data(url):
    global error_count, download_count
    driver.get(url)
    # look at insights frame
    driver.switch_to.frame(find_element(driver, 'tag_name', 'iframe', 'Could not find iframe', True))

    # wait for insights page to load
    while 'visualization' not in driver.page_source:
        time.sleep(1)
        print('waiting for page load...')
    wait_for_ajax(driver)
    time.sleep(2)

    # send shortcut to open download dialog
    body = find_element(driver, 'tag_name', 'body', 'error at body click', True)
    body.click()
    try:
        body.send_keys(Keys.CONTROL + Keys.SHIFT + 'l')
    except:
        error_count += 1
        print("error at shortcut")
    wait_for_ajax(driver)
    time.sleep(1)
    try:
        # Old method
        # driver.find_element_by_xpath(
        #    '//*[@id="lk-layout-embed"]/div[4]/div/div/form/div[2]/div[4]/div/div[2]/label').click()

        # New method of waiting for page
        waiter = wait_for_page_by_name('qr-export-modal-limit')
        waiter.click()
    except:
        print("Couldn't find 'all results' radio button. Will try again,")

    # If Appointments insight, use custom row value
    try:
        if driver.find_element_by_xpath(
                '//*[@id="lk-embed-container"]/lk-explore-dataflux/div[2]/lk-expandable-sidebar/ng-transclude/lk-field-picker/div[1]/div/div').text == "Appointments":
            driver.find_element_by_xpath(
                '//*[@id="lk-layout-embed"]/div[4]/div/div/form/div[2]/div[4]/div/div[3]/label').click()
            time.sleep(1)
            driver.find_element_by_xpath(
                '//*[@id="lk-layout-embed"]/div[4]/div/div/form/div[2]/div[4]/div/div[3]/div/div[1]/input').clear()
            time.sleep(1)
            driver.find_element_by_xpath(
                '//*[@id="lk-layout-embed"]/div[4]/div/div/form/div[2]/div[4]/div/div[3]/div/div[1]/input').send_keys(
                '99999')
    except:
        print("couldn't enter custom rows for appointments ")
    find_element(driver, 'xpath', '//*[@id="lk-layout-embed"]/div[4]/div/div/form/div[2]/div[4]/div/div[2]/label',
                 "Still couldn't find 'all results' radio button", True).click()
    print("Found 'all results' radio button.")
    time.sleep(1)
    try:
        driver.find_element_by_id('qr-export-modal-download').click()
        download_count += 1
    except:
        missed_urls[''] = url
        print(str(url))
        print("download missed")

    print('keys sent')
    time.sleep(1)


# Downloads survey results
def download_survey_data(url, driver):
    global download_count
    time.sleep(1)
    driver.get(url)
    wait_for_ajax(driver)
    find_element(driver, 'xpath', '//*[@id="ui-id-2"]/div[2]/div[1]/div/a[4]', 'could not find download button', False)
    while 'Your download is ready' not in driver.page_source:
        time.sleep(1)
    find_element(driver, 'xpath', '//*[@id="download-wait-modal"]/div[2]/div/div[2]/div/span[1]/p[1]/a',
                 'Missed survey download', True).click()
    download_count += 1
    time.sleep(60)
    driver.close()


# Downloads event information
def download_event_data(url, driver):
    global download_count
    time.sleep(1)
    driver.get(url)
    wait_for_ajax(driver)
    time.sleep(5)
    find_element(driver, 'xpath', '//*[@id="search-form"]/div/div[2]/div/div[1]/div/a/i',
                 "Could not find the download button", False)
    while 'Your download is ready' not in driver.page_source:
        time.sleep(1)
    find_element(driver, 'xpath', '//*[@id="download-wait-modal"]/div[2]/div/div[2]/div/span[1]/p[1]/a',
                 'Missed event download', True).click()
    download_count += 1
    time.sleep(60)
    driver.close()


# download_data for all URL's in dictionary. Uses multithreading to keep mulitple drivers alive and working at the
# same time.
def download_all(my_dict):
    for i in tqdm(my_dict):

        if 'insights_page' in my_dict.get(i):
            download_insight_data(my_dict[i])
        elif 'surveys' in my_dict.get(i):
            if download_file_path is not "":
                driver2 = change_download_location(download_file_path)
            else:
                driver2 = webdriver.Chrome()
            login(driver2)
            survey_thread = threading.Thread(target=download_survey_data, args=(my_dict[i], driver2,))
            survey_thread.daemon = True
            survey_thread.start()
            """ download_survey_data(my_dict[i])
            window_number += 1"""
        elif 'events' in my_dict.get(i):
            if download_file_path is not "":
                driver2 = change_download_location(download_file_path)
            else:
                driver2 = webdriver.Chrome()
            login(driver2)
            event_thread = threading.Thread(target=download_event_data, args=(my_dict[i], driver2,))
            event_thread.daemon = True
            event_thread.start()
            """download_event_data(my_dict[i])
            window_number += 1"""
        else:
            print("could not download" + str(i))

    print("Errors: " + str(error_count))

    return


# fill login form with data from config file
def login(driver):
    driver.get('https://oakland.joinhandshake.com/login')
    find_element(driver, 'xpath', '//*[@id="main"]/div[1]/div[2]/div[2]/a/div[2]', 'missed first login button',
                 True).click()
    driver.find_element_by_id('username').send_keys(username, Keys.TAB)
    driver.find_element_by_id('password').send_keys(password, Keys.ENTER)
    # find_element(driver, 'id', 'username', 'Error at username entry', True).send_keys(username, Keys.TAB)
    # find_element(driver, 'id', 'password', 'Error at password entry', True).send_keys(password, Keys.ENTER)
    while 'Student Activity Snapshot' not in driver.page_source:
        time.sleep(1)
    return


# wait for ajax to finish
def wait_for_ajax(driver):
    done = False
    while not done:
        try:
            if driver.execute_script('return jQuery.active') == 0 and driver.execute_script(
                    'return document.readyState') != 'complete':
                done = True
            print('waiting for ajax')
            time.sleep(.5)
            done = True
        except WebDriverException:
            print('WebDriverException')
    time.sleep(4)


# wait for element with name to be visible, return waiter
def wait_for_page_by_name(name):
    element = WebDriverWait().until(driver.visibility_of_element_located((By.NAME, name)))
    print(type("element type: " + element))
    return element


# wait for element with tag name to be visible, return waiter
def wait_for_page_by_tag_name(tag_name):
    element = WebDriverWait().until(driver.visibility_of_element_located((By.TAG_NAME, tag_name)))
    print(type("element type: " + element))
    return element


# wait for element with xpath to be visible, return waiter
def wait_for_page_by_xpath(xpath):
    element = WebDriverWait().until(driver.visibility_of_element_located((By.XPATH, xpath)))
    print(type("element type: " + element))
    return element


# Create main menu gui
def main_menu():
    def done():
        root.destroy()

    root = tkinter.Tk()
    root.title("Auto Downloader")
    root.minsize(280, 150)
    root.geometry("280x150")
    root.iconbitmap('updown.ico')
    Label(master=root, text="Welcome to the Handshake Automatic Downloader. \nClick start to begin.").grid(row=0,
                                                                                                           column=0)
    Label(master=root, text="").grid(row=1, column=1)

    # Create buttons.
    button1 = tkinter.Button(master=root, text='Start', command=lambda: done(), height=2, width=18)
    button2 = tkinter.Button(master=root, text='Quit', command=lambda: sys.exit(), height=2, width=18)
    button2.grid(row=4, column=0)
    button1.grid(row=3, column=0)

    root.mainloop()


# Outputs URLS for missed downloads to a file. Currently unused.
def output_missed():
    if len(missed_urls) < 1:
        return

    for i in missed_urls:
        url_str = str(i)
        print(url_str)
        print(url_str, file=open("Missed_urls_" + str(datetime.now())[:-16] + ".txt", 'a'))


# Changes drivers download location
def change_download_location(download_location):
    options = webdriver.ChromeOptions()
    options.add_experimental_option("prefs", {"download.default_directory": download_location})
    return webdriver.Chrome(executable_path='chromedriver.exe', chrome_options=options)


def find_element(driver, type, name, message, increase_error):
    time.sleep(1)
    if type.lower() == 'tag_name':
        element = ''
        try:
            element = driver.find_element_by_tag_name(name)
            return element
        except:
            if message =="Still couldn't find 'all results' radio button":
                find_element(driver, type, name, message, increase_error)
            print(message)
            if increase_error:
                error_count += 1
            return '-1'

    elif type.lower() == 'xpath':
        try:
            element =driver.find_element_by_xpath(name)
            return element
        except:
            if message == "Still couldn't find 'all results' radio button":
                find_element(driver, type, name, message, increase_error)
            print(message)
            if increase_error:
                error_count += 1
            return '-1'
    elif type.lower() == 'id':
        try:
            element = driver.find_element_by_xpath(name)
            return element
        except:
            if message == "Still couldn't find 'all results' radio button":
                find_element(driver, type, name, message, increase_error)
            print(message)
            if increase_error:
                error_count += 1
            return '-1'
    else:
        return '-1'


# Rename all of the files in the folder.
def rename_files():
    # wait for downloads.
    done_downloading = False
    while not done_downloading:
        check = glob.glob('*.crdownload')
        if len(check) <= 1:
            done_downloading = True
        else:
            print('Waiting for downloads to finish')
            time.sleep(1)
    time.sleep(30)

    os.chdir(download_file_path)
    files = glob.glob('generated*.csv')
    print('Renaming Files')
    for file in files:
        filename = os.path.basename(file)[31:-19] + datetime.now().strftime("%Y-%m-%d_%H.%M.%S") + ".csv"
        os.rename(file, filename)
        time.sleep(1)
    files = glob.glob('event_download*.csv')
    for file in files:
        filename = os.path.basename(file)[:5] + "_results" + datetime.now().strftime("%Y-%m-%d_%H.%M.%S") + ".csv"
        os.rename(file, filename)
        time.sleep(1)
    files = glob.glob('survey_response*.csv')
    for file in files:
        filename = os.path.basename(file)[:16] + datetime.now().strftime("%Y-%m-%d_%H.%M.%S") + ".csv"
        os.rename(file, filename)
        time.sleep(1)


# Initialize variables and being download.
def main():
    global driver, input_file_path, number_of_rows, download_count, missed_urls
    download_count = 0
    input_file_path = ""
    number_of_rows = 0
    load_config()
    if download_file_path is not "":
        driver.close()
        driver = change_download_location(download_file_path)
    login(driver)
    csv = load_csv()
    download_all(csv)
    if len(missed_urls) > 0:
        download_all(missed_urls)
    time.sleep(60)
    rename_files()
    print(str(download_count) + " of " + str(len(csv)) + " files downloaded"+"\nFiles saved to "+str(download_file_path))



# define and initialize global variables
driver = webdriver.Chrome()
error_count = 0
print('error count: ' + str(error_count))
download_count = 0
input_file_path = ""
download_file_path = ""
number_of_rows = 0
missed_urls = {}
main()
