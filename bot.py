# Imports
import os, sys, shutil, datetime, re, time, csv
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

# Action booleans
download_leads_file = True  # Production value: True
upload_leads_to_bt = True  # Production value: True
reset_google_sheets = True  # Production value: True
pause_on_error = False  # Production value: False
send_status_email = True  # Production value: True


def clear():
    os.system('cls' if os.name == 'nt' else 'clear')

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
        os.remove(get_config_data(DataCategories.LEADS_DATA, 'leads_download_filepath'))
    except Exception as e:
        print(str(e))
    # Download new leads file
    driver.get(get_config_data(DataCategories.LEADS_DATA, 'leads_file_download_url'))
    i = 0
    num_itrs = 1000
    for i in range(num_itrs):
        clear()
        print('Waiting for download...')
        if os.path.exists(get_config_data(DataCategories.LEADS_DATA, 'leads_download_filepath')):
            break
    if i == num_itrs-1:
        os.system('cls' if os.name == 'nt' else 'clear')
        print('Leads file failed to download.')
        closeChrome()
    else:
        print('Download successful!')
    # Copy leads file to archive location. Append timestamp to filename
    timestamp = '_' + str(datetime.datetime.now()).replace('-','').replace(' ','_').replace(':','')
    timestamp = timestamp[:timestamp.find('.')]
    try:
        src_filepath = get_config_data(DataCategories.LEADS_DATA, 'leads_download_filepath')
        filename = os.path.basename(src_filepath)
        dst_filepath = get_config_data(DataCategories.LEADS_DATA, 'leads_archive_directory') + '\\' +filename[:filename.find('.')] + timestamp + '.csv'
        shutil.copyfile(src_filepath, dst_filepath)
    except Exception as e:
        print('-- ERROR --')
        print(str(e))
        if send_status_email: newStatus('Could not copy leads download file to archive folder: ' + str(e), True)
        if pause_on_error: input('Press [Enter]... ')
        closeChrome()
    # Check if the csv is blank
    with open(get_config_data(DataCategories.LEADS_DATA, 'leads_download_filepath')) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        line_count = 0
        col_qty = 0
        blank_csv = True
        for row in csv_reader:
            if line_count == 0:
                for colname in row:
                    if len(colname) != 0:
                        col_qty += 1
                line_count += 1
            else:
                for i in range(len(row)):
                    if i >= 16:
                        break 
                    if row[i] != '' and row[i] != ' ':
                        blank_csv = False
                        break
                line_count += 1
    if blank_csv:
        no_leads_msg = 'Downloaded Leads csv was empty. No leads to upload.'
        print(no_leads_msg)
        if send_status_email: newStatus(no_leads_msg, False)
        if pause_on_error: input('Press [Enter]... ')
        closeChrome()


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
    clear()
    print('Attempting leads upload')
    # Click import leads button
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="rc-tabs-0-panel-ListView"]/header/a[2]/button'))).click()
    # Select file to upload
    file_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.ImportWizard .UploadButton input')))
    file_input.send_keys(get_config_data(DataCategories.LEADS_DATA, 'leads_download_filepath'))
    # Click 'Next' button until the results page reached
    tries = 20
    itr = 0
    while itr < tries:
        # Check if error message is displayed. If so, send status err message and close program
        try:
            error_container = driver.find_element(By.CSS_SELECTOR, 'div.ant-alert-error')
        except NoSuchElementException:
            pass
        else:
            err_msg = error_container.find_element(By.CSS_SELECTOR, 'div.ant-alert-description > ul > li').text
            print('ERROR OCCURED:', err_msg)
            if send_status_email: newStatus(err_msg, True)
            if pause_on_error: input('Press [Enter]... ')
            closeChrome()
        # Attempt to click "Next" button
        try:
            itr += 1
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.ImportWizard [data-testid="next"]'))).click()
        except (StaleElementReferenceException, ElementClickInterceptedException): 
            try:
                # Check for import wizard results container
                driver.find_element(By.CSS_SELECTOR, '.ImportWizard .importWizardResults')
            except NoSuchElementException:
                pass
            else:
                break
    # Check if import was successful
    try:
        WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.ImportWizard .success-message')))
    except (NoSuchElementException, TimeoutException):
        err_msg = 'Upload was NOT successful. No element could be found matching ".ImportWizard .success-message"'
        print(err_msg)
        if send_status_email: newStatus(err_msg, True)
        if pause_on_error: input('Press [Enter]... ')
        closeChrome()
    else:
        print('Upload was successful!')


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
        if download_leads_file:
            downloadLeadsFile()
        if upload_leads_to_bt:
            btLogin()
            btUploadLeads()
        if reset_google_sheets:
            googleLogin()
            clearSheets()
        if send_status_email: newStatus('Program ran successfully', False)
        closeChrome()
    except Exception as e:
        print('Error encountered')
        if send_status_email: newStatus('ERROR encountered while running program:\n ' + str(e), True)
        if pause_on_error: input('Press [Enter]... ')
        sys.exit()
