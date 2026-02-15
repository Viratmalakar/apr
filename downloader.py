from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time
import os

def download_reports():

    download_folder = os.path.expanduser("~/Downloads")

    chrome_options = Options()

    prefs = {
        "download.default_directory": download_folder,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True
    }

    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--start-maximized")

    driver = webdriver.Chrome(options=chrome_options)

    # OPEN PORTAL
    driver.get("http://192.168.2.222:8080/reportsngucc/Home")

    time.sleep(3)

    # LOGIN
    driver.find_element(By.ID, "username").send_keys("supervisor@csr")
    driver.find_element(By.ID, "password").send_keys("bfil@000")

    driver.find_element(By.ID, "loginBtn").click()

    time.sleep(5)


    # CLICK REPORTS
    driver.find_element(By.LINK_TEXT, "Reports").click()
    time.sleep(2)

    driver.find_element(By.LINK_TEXT, "Agent Performance").click()
    time.sleep(3)


    # SELECT ALL AGENTS
    driver.find_element(By.ID, "agentSelectAll").click()

    time.sleep(2)


    # CLICK SUBMIT
    driver.find_element(By.ID, "submitBtn").click()

    time.sleep(5)


    # CLICK EXCEL DOWNLOAD
    driver.find_element(By.ID, "excelBtn").click()

    time.sleep(10)


    driver.quit()

