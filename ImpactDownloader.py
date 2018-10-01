# Impact Downloader script by Nick Furlo 5/21/18
import configparser
import csv
import fnmatch
import glob
import itertools
import os
import shutil
import threading
import time
import tkinter
from datetime import datetime
from distutils.util import strtobool
from os import listdir
from os.path import isfile, join
from pathlib import Path
from tkinter import *
from tkinter import messagebox

from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
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
            "[DEFAULT]\nUSERNAME = \nPASSWORD = \nINPUT_CSV_FILE_PATH = \nNUMBER_OF_ROWS = \nDOWNLOAD_LOCATION = \nNETWORK_LOCATION =  \nLOG_TO_FILE = \nDELETE_AFTER_DAYS = ")
        messagebox.showinfo("Warning", "Config file created. Please add a CSV file path and relaunch the program.")
        driver.close()
        sys.exit()

    # Read config and set variables
    config.read('ImpactDownloader.config')
    global username, password, input_file_path, number_of_rows, download_file_path, network_location, log_enabled, days_until_delete
    username = config['DEFAULT']['USERNAME']
    password = config['DEFAULT']['PASSWORD']
    input_file_path = config['DEFAULT']['INPUT_CSV_FILE_PATH']
    number_of_rows = int(config['DEFAULT']['NUMBER_OF_ROWS'])
    download_file_path = config['DEFAULT']['DOWNLOAD_LOCATION']
    network_location = config['DEFAULT']['NETWORK_LOCATION']
    log_enabled = config['DEFAULT']['LOG_TO_FILE']
    days_until_delete = config['DEFAULT']['DELETE_AFTER_DAYS']

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
def download_insight_data(url, folder):
    global error_count, download_count
    driver.get(url)
    # look at insights frame
    driver.switch_to.frame(find_element(driver, 'tag_name', 'iframe', 'Could not find iframe', True, False, False))

    # wait for insights page to load
    while 'visualization' not in driver.page_source:
        time.sleep(1)
        print('waiting for page load...')
    time.sleep(2)

    # send shortcut to open download dialog
    wait_for_page('tag_name', 'body')
    body = find_element(driver, 'tag_name', 'body', 'error at body click', True, False, False)
    body.click()
    try:
        while("With visualization options applied " not in driver.page_source):
            time.sleep(2)
            body.click()
            body.send_keys(Keys.CONTROL + Keys.SHIFT + 'l')
            print("Attempting Shortcut")
        print("Sent shortcut")
    except:
        error_count += 1
        print("error at shortcut")
    time.sleep(1)
    try:
        # Old method
        # driver.find_element_by_xpath(
        #    '//*[@id="lk-layout-embed"]/div[4]/div/div/form/div[2]/div[4]/div/div[2]/label').click()

        # New method of waiting for page
        wait_for_page('name', 'qr-export-modal-limit')
        find_element(driver, 'xpath', '//*[@id="lk-layout-embed"]/div[4]/div/div/form/div[2]/div[4]/div/div[2]/label',
                     "Couldn't find 'all results' radio button", True, False, True)
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
                 "Still couldn't find 'all results' radio button", True, False, True)
    print("Found 'all results' radio button.")
    time.sleep(1)

    try:
        # Change filename
        wait_for_page('xpath', '//*[@id="qr-export-modal-custom-filename"]')
        filename = folder + datetime.now().strftime("_%Y-%m-%d") + ".csv"
        driver.find_element_by_xpath('//*[@id="qr-export-modal-custom-filename"]').clear()
        driver.find_element_by_xpath('//*[@id="qr-export-modal-custom-filename"]').send_keys(filename)
    except:
        print("Could not rename file before downloading.")
        error_count += 1

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
    driver.find_element_by_xpath('//*[@id="ui-id-2"]/div[2]/div[1]/div/a[4]').click()
    while 'Your download is ready' not in driver.page_source:
        time.sleep(1)
    driver.find_element_by_xpath('//*[@id="download-wait-modal"]/div[2]/div/div[2]/div/span[1]/p[1]/a').click()
    os.chdir(download_file_path)
    while next(glob.iglob('*.crdownload'), False) is not False:
        time.sleep(1)
    print('survey data done downloading')
    download_count += 1
    files = glob.glob('survey_response_download*.csv')
    for file in files:
        os.rename(file, "appt_survey" + datetime.now().strftime("_%Y-%m-%d") + ".csv")
    driver.close()


# Downloads event information
def download_event_data(url, driver):
    global download_count
    time.sleep(1)
    driver.get(url)
    time.sleep(5)
    driver.find_element_by_xpath('//*[@id="search-form"]/div/div[2]/div/div[1]/div/a').click()
    while 'Your download is ready' not in driver.page_source:
        time.sleep(1)
    driver.find_element_by_xpath('//*[@id="download-wait-modal"]/div[2]/div/div[2]/div/span[1]/p[1]/a').click()
    os.chdir(download_file_path)
    while next(glob.iglob('*.crdownload'), False) is not False:
        time.sleep(1)
    print('event data done downloading')
    download_count += 1
    files = glob.glob('event_download*.csv')
    for file in files:
        os.rename(file, "event_detail" + datetime.now().strftime("_%Y-%m-%d") + ".csv")
    driver.close()


# download_data for all URL's in dictionary. Uses multithreading to keep mulitple drivers alive and working at the
# same time.
def download_all(my_dict):
    for i in tqdm(my_dict):

        if 'insights_page' in my_dict.get(i):
            my_dict_reverse = {v: k for k, v in my_dict.items()}
            download_insight_data(my_dict[i], my_dict_reverse[my_dict[i]])
        elif 'surveys' in my_dict.get(i):
            if download_file_path is not "":
                driver2 = change_download_location(download_file_path)
            else:
                driver2 = webdriver.Chrome()
            login(driver2)
            survey_thread = threading.Thread(target=download_survey_data, args=(my_dict[i], driver2,))
            survey_thread.daemon = True
            survey_thread.start()

        elif 'events' in my_dict.get(i):
            if download_file_path is not "":
                driver2 = change_download_location(download_file_path)
            else:
                driver2 = webdriver.Chrome()
            login(driver2)
            event_thread = threading.Thread(target=download_event_data, args=(my_dict[i], driver2,))
            event_thread.daemon = True
            event_thread.start()

        else:
            print("could not download" + str(i))

    print("Errors: " + str(error_count))

    return


# fill login form with data from config file
def login(driver):
    driver.get('https://oakland.joinhandshake.com/login')
    driver.find_element_by_xpath('//*[@id="main"]/div[1]/div[2]/div[2]/a/div[2]').click()
    # find_element(driver, 'xpath', '//*[@id="main"]/div[1]/div[2]/div[2]/a/div[2]', 'Missed body on first login page', False, True, True)
    driver.find_element_by_id('username').send_keys(username, Keys.TAB)
    driver.find_element_by_id('password').send_keys(password, Keys.ENTER)
    while 'Student Activity Snapshot' not in driver.page_source:
        time.sleep(1)
    return


# Called by find_element() and used to wait until a specific element is loaded.
def wait_for_page(type, name):
    if type.lower() == 'name':
        # print("Presence of element: " + str(driver.find_element_by_name(name).is_displayed()))
        # WebDriverWait().until(driver, driver.find_element_by_name(name).is_displayed())
        WebDriverWait(driver, 900).until(EC.visibility_of_element_located((By.NAME, name)))
        return  # print('wait_for_page success')

    elif type.lower() == 'xpath':
        # print("Presence of element: " + str(driver.find_element_by_xpath(name).is_displayed()))
        # WebDriverWait().until(driver, driver.find_element_by_xpath(name).is_displayed())
        WebDriverWait(driver, 900).until(EC.visibility_of_element_located((By.XPATH, name)))
        return  # print('wait_for_page success')

    elif type.lower() == 'tag_name':
        # print("Presence of element: " + str(driver.find_element_by_tag_name(name).is_displayed()))
        # WebDriverWait().until(driver, driver.find_element_by_tag_name(name).is_displayed())
        WebDriverWait(driver, 900).until(EC.visibility_of_element_located((By.TAG_NAME, name)))
        return  # print('wait_for_page success')

    elif type.lower() == 'id':
        # print("Presence of element: " + str(driver.find_element_by_id(name).is_displayed()))
        # WebDriverWait().until(driver, driver.find_element_by_id(name).is_displayed())
        WebDriverWait(driver, 900).until(EC.presence_of_element_located((By.ID, name)))
        return  # print('wait_for_page success')
    elif type.lower() == 'class_name':
        # print("Presence of element: " + str(driver.find_element_by_class_name(name).is_displayed()))
        # WebDriverWait().until(driver, driver.find_element_by_id(name).is_displayed())
        WebDriverWait(driver, 900).until(EC.presence_of_element_located((By.class_name, name)))
        return  # print('wait_for_page success')
    else:
        return -'1'


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


# Runs driver.find_element_by_TYPE(name) surrounded by error handling.
# driver is driver to be used, type is the type of element, name is the
# element location data, message is error to be printed to console,
# increase error is boolean that decides to increment error counter
def find_element(driver, type, name, message, increase_error, recurse, click):
    time.sleep(1)
    if type.lower() == 'tag_name':
        try:
            wait_for_page('tag_name', name)
            if click:
                return driver.find_element_by_tag_name(name).click()
            return driver.find_element_by_tag_name(name)
        except:
            print(message)
            if recurse:
                find_element(driver, type, name, message, increase_error, recurse, click)
            if increase_error:
                error_count += 1
            return find_element(driver, type, name, message, increase_error, recurse, click)
    elif type.lower() == 'name':
        try:
            wait_for_page('name', name)
            if click:
                return driver.find_element_by_name(name).click()
            return driver.find_element_by_name(name)
        except:
            print(message)
            if recurse:
                find_element(driver, type, name, message, increase_error, recurse, click)
            if increase_error:
                error_count += 1
            return find_element(driver, type, name, message, increase_error, recurse, click)

    elif type.lower() == 'xpath':
        try:
            print('looking for xpath')
            wait_for_page('xpath', name)
            if click:
                return driver.find_element_by_xpath(name).click()
            return driver.find_element_by_xpath(name)

        except:
            print(message)
            if recurse:
                find_element(driver, type, name, message, increase_error, recurse, click)
            if increase_error:
                error_count += 1
            return find_element(driver, type, name, message, increase_error, recurse, click)
    elif type.lower() == 'id':
        try:
            wait_for_page('id', name)
            if click:
                return driver.find_element_by_id(name).click()
            return driver.find_element_by_id(name)

        except:
            print(message)
            if recurse:
                find_element(driver, type, name, message, increase_error, recurse, click)
            if increase_error:
                error_count += 1
            return find_element(driver, type, name, message, increase_error, recurse, click)
    elif type.lower() == 'class_name':
        try:
            wait_for_page('id', name)
            if click:
                return driver.find_element_by_class_name(name).click()
            return driver.find_element_by_class_name(name)

        except:
            print(message)
            if recurse:
                find_element(driver, type, name, message, increase_error, recurse, click)
            if increase_error:
                error_count += 1
            return find_element(driver, type, name, message, increase_error, recurse, click)
    else:
        return '-1'


# Wait for downloads to finish.
def wait_for_downloads():
    # wait for downloads.
    os.chdir(download_file_path)
    print("Waiting for downloads to finish")
    wait_count = 0
    while len(glob.glob('*.crdownload')) > 0:
        print("Waiting for downloads to finish...")
        time.sleep(6)
        wait_count += 1
        if wait_count > 1000:
            time.sleep(5)
            break
    time.sleep(5)


# Copies files into folder with names the same as the name of the file without datetime. These folders are in a
# network location specified by the user.
def copy_to_network_drive():
    copied = 0
    try:
        print("Starting copy to network location:" + str(network_location))
        os.chdir(download_file_path)
        files = glob.glob("*.csv")
    except:
        print("Could not get files to move to network location.")
    try:
        for file in files:
            full_path = os.path.join(network_location, file[:-15])
            full_path = full_path.replace("\\", "/")
            network_paths.append((full_path))
            shutil.copy2(file, full_path)
            print("Copied: " + full_path)
            copied += 1
    except Exception as e:
        print("Could not copy files to network location")
        print("Error code: " + str(e))
    print(str(copied) + " files coppied to " + str(network_location))


def delete_csv_from_download():
    fileList = os.listdir(download_file_path)
    for fileName in fileList:
        try:
            os.remove(download_file_path + "/" + fileName)
        except Exception as e:
            print("Could not delete " + str(fileName) + " because: " + str(e))
    print("CSV files deleted from download directory")


def log_to_file():
    path = str(os.getcwd() + '/Logs').replace('\\', '/')
    if not os.path.exists(path):
        os.mkdir(path)

    full_path = path + "/" + "ImpactDownloaderLog" + datetime.now().strftime("_%Y-%m-%d_%H.%M") + ".txt"
    sys.stdout = open(full_path, "w")


# Deletes downloaded files from the network locations that are more than 7 days old.
# Substrings file name, creates datetime from substring, compares datetime, deletes or not.
def delete_old_network_files(cutoff):
    now = time.time()
    print("Checking for files older than " + str(cutoff) + " days.")

    for path in network_paths:

        files = os.listdir(path)
        print(str(files))
        for xfile in files:
            full_pth = Path(os.path.join(path, xfile))
            if full_pth.exists():
                try:
                    created = str(full_pth.name)[-14:-4]
                    created_datetime = datetime.strptime(created, "%Y-%m-%d")
                    print(str(xfile) + " created: " + str(created_datetime))
                    today = datetime.now().strftime("%Y-%m-%d")
                    today_datetime = datetime.strptime(today, "%Y-%m-%d")
                    difference = today_datetime - created_datetime
                    # delete file if older than a week
                    if int(difference.total_seconds() / 86400) > cutoff:
                        print(str(xfile) + " is older than " + str(cutoff) + " days. Deleting File...")
                        os.remove(str(full_pth))
                        print(str(xfile) + " was deleted.")
                    else:
                        print("File is not more than " + str(cutoff) + " days old, not deleted.")
                except Exception as e:
                    print(str(xfile) + " was not deleted.")
                    print("error code: " + str(e))
            else:
                print("No old files found")


# Initialize variables and being download.
def main():
    delete_csv_from_download()
    global driver, input_file_path, number_of_rows, download_count, missed_urls, log_to_file
    download_count = 0
    input_file_path = ""
    number_of_rows = 0
    load_config()
    if bool(strtobool(str(log_enabled))):
        log_to_file()
    if download_file_path is not "":
        driver.close()
        driver = change_download_location(download_file_path)
    login(driver)
    csv = load_csv()
    download_all(csv)
    if len(missed_urls) > 0:
        download_all(missed_urls)
    wait_for_downloads()
    if network_location != "":
        copy_to_network_drive()
    delete_csv_from_download()
    if int(days_until_delete) > 0:
        delete_old_network_files(int(days_until_delete))
        print("Old Files Deleted")
    driver.close()
    print(str(download_count) + " of " + str(number_of_rows - 1) + " files downloaded" + "\nFiles saved to " + str(
        download_file_path) + " and moved to " + str(network_location))
    print("\n\nImpact Downloader Finished")


# define and initialize global variables
driver = webdriver.Chrome()
error_count = 0
download_count = 0
input_file_path = ""
download_file_path = ""
network_location = ""
number_of_rows = 0
log_enabled = ""
missed_urls = {}
global network_paths
network_paths = []
days_until_delete = 0
main()
