# Imports
import os, sys, shutil, datetime, re, time
import undetected_chromedriver as uc
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException, ElementClickInterceptedException, TimeoutException, SessionNotCreatedException, WebDriverException
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from win32com.client import Dispatch

from config import DataCategories, get_config_data
from bot_status import newStatus


# Initialize globals
driver = None
chrome_version = None
chrome_options = None
actionChains = None

MIN_CHROME_VERSION = 102 # What chromedriver version to start at

# Browser to be used
BROWSER = 'Undetected Chrome'  # Options: 'Chrome', 'Undetected Chrome', 'Firefox'


# Find the current chrome version
def get_version_via_com(filename):
    parser = Dispatch("Scripting.FileSystemObject")
    try:
        version = parser.GetFileVersion(filename)
    except Exception:
        return None
    return version

# Set the chrome version global
def set_chrome_version_global():
    paths = [r"C:\Program Files\Google\Chrome\Application\chrome.exe",
             r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
             r"C:\Users\jordan\AppData\Local\Google\Chrome\Application\chrome.exe"]
    version = list(filter(None, [get_version_via_com(p) for p in paths]))[0]
    version = version[:version.find('.')]  # Grab just the first number
    # Set to global
    global chrome_version
    chrome_version = version

# Initialize chrome driver
def initDriver():
    # Declare driver
    global driver
    # Chrome
    if BROWSER == 'Chrome' or BROWSER == 'Undetected Chrome':
        set_chrome_version_global()
        directory = get_config_data(DataCategories.CHROMEDRIVER_DATA, 'directory').replace('/', '\\')
        chromedriver_path = (os.path.realpath(__file__)[::-1][(os.path.realpath(__file__)[::-1].find('\\')+1):])[::-1] + directory + 'chromedriver_' + str(chrome_version) + '_win.exe'
        if BROWSER == 'Chrome':
            driver = webdriver.Chrome(executable_path=chromedriver_path)
        elif BROWSER == 'Undetected Chrome':
            driver = uc.Chrome(executable_path=chromedriver_path)
    # Firefox
    elif BROWSER == 'Firefox':
        try:
            profile = webdriver.FirefoxProfile(get_config_data(DataCategories.FIREFOXDRIVER_DATA, 'profile_path'))
            profile.set_preference('dom.webdriver.enabled', False)
            profile.set_preference('useAutomationExtension', False)
            profile.update_preferences()
            desired = DesiredCapabilities.FIREFOX
            driver = webdriver.Firefox(firefox_profile=profile, desired_capabilities=desired)
        except Exception as e:
            print(str(e))
    else:
        raise ValueError('Not set up for this browser yet (' + BROWSER + ')')
    driver.maximize_window()


# Initialize actionChains. This is used to perform actions like scroling elements into view
def initActionChains():
    global actionChains
    actionChains = ActionChains(driver)


# Download leads file using link
def downloadLeadsFile():
    # Delete leads file before new download
    try:
        os.remove(get_config_data(DataCategories.LEADS_DATA, 'leads_filepath'))
    except Exception as e:
        print(str(e))
    # Download new leads file
    driver.get(get_config_data(DataCategories.LEADS_DATA, 'leads_file_download_url'))
    i = 0
    num_itrs = 1000
    for i in range(num_itrs):
        os.system('cls' if os.name == 'nt' else 'clear')
        print('Checking for download...')
        if os.path.exists(get_config_data(DataCategories.LEADS_DATA, 'leads_filepath')):
            break
    if i == num_itrs-1:
        os.system('cls' if os.name == 'nt' else 'clear')
        print('Leads file failed to download.')
        input('Press [ENTER] to close program ')
        sys.exit()
    # Copy leads file to archive location. Append timestamp to filename
    timestamp = '_' + str(datetime.datetime.now()).replace('-','').replace(' ','_').replace(':','')
    timestamp = timestamp[:timestamp.find('.')]
    try:
        src_filepath = get_config_data(DataCategories.LEADS_DATA, 'leads_filepath')
        filename = os.path.basename(src_filepath)
        dst_filepath = get_config_data(DataCategories.LEADS_DATA, 'leads_archive_directory') + '\\' +filename[:filename.find('.')] + timestamp + '.csv'
        shutil.copyfile(src_filepath, dst_filepath)
    except Exception as e:
        print('-- ERROR --')
        print(str(e))
        input('Press [ENTER] to close program ')
        sys.exit()


# Log into BuilderTrend
def btLogin():
    # Go to the login url
    driver.get(get_config_data(DataCategories.BT_LOGIN_DATA, 'login_url'))
    # Log in
    driver.find_element(By.XPATH, '//*[@id="username"]').send_keys(get_config_data(DataCategories.BT_LOGIN_DATA, 'username'))  # Enter username
    driver.find_element(By.XPATH, '//*[@id="password"]').send_keys(get_config_data(DataCategories.BT_LOGIN_DATA, 'password'))  # Enter password
    # Click login button
    login_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#reactLoginListDiv button')))
    actionChains.move_to_element(login_btn).click().perform()


# Upload leads file to BuilderTrend
def btUploadLeads():
    # Click import leads button
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="rc-tabs-0-panel-ListView"]/header/a[2]/button'))).click()
    # Select file to upload
    file_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.ImportWizard .UploadButton input')))
    file_input.send_keys(get_config_data(DataCategories.LEADS_DATA, 'leads_filepath'))
    # Click 'Next' button until the results page reached
    while True:
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.ImportWizard [data-testid="next"]'))).click()
        except (StaleElementReferenceException, ElementClickInterceptedException): 
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
        # newStatus('Import was not successful!', True)
        print('Import was NOT successful')
        input('Press [ENTER] ')
        sys.exit()
    else:
        # newStatus('Import was successful!', False)
        print('Import was successful!')


# Login to EH gmail account
def googleLogin():
    try:
        driver.get('https://accounts.google.com/')
        # enter email and password
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#identifierId'))).send_keys(get_config_data(DataCategories.EH_GMAIL_LOGIN_DATA, 'email'))
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#identifierNext button'))).click()
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#password input'))).send_keys(get_config_data(DataCategories.EH_GMAIL_LOGIN_DATA, 'password'))
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#passwordNext button'))).click()
    except Exception as e:
        print('ERROR: Could not log into google account')
        raise e


# clicks the clear/reset buttons in the Lead google sheets
def clearSheets():
    ## Lead Sheet 1
    # Open new lead sheet 1
    new_lead_sheet_1_url = get_config_data(DataCategories.LEADS_DATA, 'new_lead_sheet_1_url')
    while not re.search(new_lead_sheet_1_url+'.*',driver.current_url):
        print('trying url...')
        driver.get(new_lead_sheet_1_url)
    print('url reached')
    # Click 'Clear Sheet' button
    while True:
        print('attempting click')
        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#fixed_right .waffle-borderless-embedded-object-container div:nth-child(1)'))).click()
        except ElementClickInterceptedException:
            try:
                driver.find_element(By.CSS_SELECTOR, 'div.docs-error-dialog button').click()
            except NoSuchElementException:
                print('Could not click Clear Sheet button on the first sheet, something is blocking it and it is not the sign-in dialog.')
                input('Press [ENTER] to close program ')
                closeChrome()
        else:
            try:
                WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//div[@id="docs-butterbar-container"]/div/div[@role="alert" and text()="Finished script"]')))
                break
            except TimeoutException:
                pass
    
    ## Lead Sheet 2
    # Open new lead sheet 2
    new_lead_sheet_2_url = get_config_data(DataCategories.LEADS_DATA, 'new_lead_sheet_2_url')
    while not re.search(new_lead_sheet_2_url+'.*',driver.current_url):
        print('trying url...')
        driver.get(new_lead_sheet_2_url)
    print('url reached')
    # Click 'Clear Sheet' button
    while True:
        print('attempting click')
        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#fixed_right .waffle-borderless-embedded-object-container div:nth-child(1)'))).click()
        except ElementClickInterceptedException:
            try:
                driver.find_element(By.CSS_SELECTOR, 'div.docs-error-dialog button').click()
            except NoSuchElementException:
                print('Could not click Clear Sheet button on the second sheet, something is blocking it and it is not the sign-in dialog.')
                input('Press [ENTER] to close program ')
                closeChrome()
        else:
            try:
                WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//div[@id="docs-butterbar-container"]/div/div[@role="alert" and text()="Finished script"]')))
                break
            except TimeoutException:
                pass


    ## Lead Sheet 3
    # Open new lead sheet 3
    new_lead_sheet_3_url = get_config_data(DataCategories.LEADS_DATA, 'bt_lead_upload_sheet_url')
    while not re.search(new_lead_sheet_3_url+'.*',driver.current_url):
        print('trying url...')
        driver.get(new_lead_sheet_3_url)
    print('url reached')
    # Click 'ResetSheet' dropdown
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#docs-extensions-menu'))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'span[aria-label="Macros m"]'))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//span[text()="ResetSheet"]'))).click()
    time.sleep(2)


# Stop the program
def stop():  # TODO: remove this and all instances after dev
    print('Program stopped')
    input('Press [ENTER] to close program ')
    closeChrome()


# Close chrome driver bot
def closeChrome():
    os.system('cls' if os.name == 'nt' else 'clear')  #TODO: Uncomment when dev finished
    print('Closing driver...')
    driver.quit()
    print('\n\nDriver Closed\n')
    sys.exit()


# Main method
if __name__ == '__main__':
    try:
        initDriver()
        initActionChains()
        downloadLeadsFile()
        btLogin()
        btUploadLeads()
        googleLogin()
        clearSheets()
        closeChrome()
    except Exception as e:
        # newStatus('Exception raised while running program:\n' + str(e), True)
        print('Error encountered')
        input('Press [ENTER] to view')
        raise e
    sys.exit()
