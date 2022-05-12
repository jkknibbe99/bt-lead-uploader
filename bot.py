# Imports
from config import DataCategories, get_config_data #login_data, chromedriver_data
from bot_status import newStatus
import os, sys, time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import WebDriverException, NoSuchElementException, ElementClickInterceptedException, StaleElementReferenceException, TimeoutException


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
    # Select file to upload
    file_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.ImportWizard .UploadButton input')))
    file_input.send_keys(get_config_data(DataCategories.INIT_DATA, 'leads_file_path'))
    # Click 'Next' button until the results page reached
    while True:
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.ImportWizard [data-testid="next"]'))).click()
        except StaleElementReferenceException:
            try:
                driver.find_element(By.CSS_SELECTOR, '.ImportWizard .importWizardResults')
            except NoSuchElementException:
                pass
            else:
                break
    # Check if import was successful
    try:
        WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.ImportWizard .success-message')))
    except NoSuchElementException:
        newStatus('Import was not successful!', True)
    else:
        newStatus('Import was successful!', False)


# Close chrome driver bot
def closeChrome():
    os.system('cls' if os.name == 'nt' else 'clear')
    print('\n\nChromedriver Closed\n')
    driver.quit()


# Main method
if __name__ == '__main__':
    try:
        initDriver()
        initActionChains()
        login()
        importLeads()
        time.sleep(2)
        closeChrome()
    except Exception as e:
        newStatus('Exception raised while running program:\n' + str(e), True)
    sys.exit()
