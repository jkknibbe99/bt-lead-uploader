# Imports
from config import DataCategories, get_config_data #login_data, chromedriver_data
from bot_status import newStatus
import os, sys, time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains


# Initialize globals
driver = None
chrome_options = None
actionChains = None


# Initialize chrome driver
def initDriver():
    # Get path to chromedriver
    if os.name == 'nt':  # Windows
        directory = get_config_data(DataCategories.CHROMEDRIVER_DATA, 'directory').replace('/', '\\')
        chromedriver_path = (os.path.realpath(__file__)[::-1][(os.path.realpath(__file__)[::-1].find('\\')+1):])[::-1] + directory + 'chromedriver_100_win.exe'
        print(chromedriver_path)
    else:  # Mac
        raise NotImplementedError
    # Declare chromedriver
    global driver
    if chrome_options is None:
        driver = webdriver.Chrome(executable_path=chromedriver_path)
    else:
        driver = webdriver.Chrome(executable_path=chromedriver_path,options=chrome_options)


# Initialize actionChains. This is used to perform actions like scroling elements into view
def initActionChains():
    global actionChains
    actionChains = ActionChains(driver)


# Log into BuilderTrend
def login():
    # Go to the login url
    driver.get(get_config_data(DataCategories.BT_LOGIN_DATA, 'login_url'))
    # Log in
    driver.find_element(By.XPATH, '//*[@id="username"]').send_keys(get_config_data(DataCategories.BT_LOGIN_DATA, 'username'))  # Enter username
    driver.find_element(By.XPATH, '//*[@id="password"]').send_keys(get_config_data(DataCategories.BT_LOGIN_DATA, 'password'))  # Enter password
    # Click login button
    login_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#reactLoginListDiv button')))
    actionChains.move_to_element(login_btn).click().perform()


def importLeads():
    # Click import leads button
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="rc-tabs-0-panel-ListView"]/header/a[2]/button'))).click()
    
    file_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.ImportWizard .UploadButton input')))
    file_input.send_keys(get_config_data(DataCategories.INIT_DATA, 'leads_file_path'))


# Request a bulk export
def requestBulkExport():
    # Click Settings dropdown button
    settings_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="reactMainNavigation"]/div[1]/div/ul/li[12]/div')))
    actionChains.move_to_element(settings_btn).click().perform()
    # Find and click Setup dropdown option
    dropdown_div = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//div[@id="reactMainNavigation"]/div[1]/div[1]/div')))
    setup_elem = dropdown_div.find_element(By.XPATH, '//div/div/ul/li[1]/span/div/a/div/div[text()="Setup"]')
    setup_elem = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(setup_elem))
    actionChains.move_to_element(setup_elem).click().perform()
    # Switch to Setup iframe
    iframes = driver.find_elements(By.TAG_NAME, 'iframe')
    for iframe in iframes:
        if iframe.get_dom_attribute('title') == 'Setup':
            setup_iframe = iframe
            break
    driver.switch_to.frame(setup_iframe)
    # Click Bulk Export button
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="root"]/div[1]/section/main/div/div[2]/div[1]/div[2]/div/div[8]/a/div'))).click()
    # Select Open and Closed Jobs
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="root"]/div[1]/section/main/div[2]/div[1]/div[1]/div/div/div[2]/form/div[1]/div/div/label[2]/span[1]'))).click()
    # Click on the Request Export button
    export_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="root"]/div[1]/section/main/div[2]/div[1]/div[1]/div/div/div[2]/form/div[3]/div/div[1]/div/button/span')))
    if export_btn.text == 'Cancel Export':
        newStatus('Attempted to request a bulk export, however, it looks like one was already requested.', True)
    else:
        actionChains.move_to_element(export_btn).click().perform()
        newStatus('Bulk export requested successfully', False)


# Close chrome driver bot
def closeChrome():
    os.system('cls' if os.name == 'nt' else 'clear')
    print('\n\nChromedriver Closed\n')
    driver.quit()


# Main method
if __name__ == '__main__':
    initDriver()
    initActionChains()
    login()
    importLeads()
    time.sleep(2)
    closeChrome()
    sys.exit()
