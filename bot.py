# Imports
import os, sys, shutil, datetime, re, time, csv
import undetected_chromedriver as uc
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException, ElementClickInterceptedException, TimeoutException, WebDriverException
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from win32com.client import Dispatch
from config import DataCategories, get_config_data
from bot_status import newStatus
from send_email import sendEmail
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd
from chromedrivers.install_latest_chromedriver import installLatestChromedriver
import chromedriver_autoinstaller


# Initialize globals
driver = None
chrome_version = None
chrome_options = None
actionChains = None

MIN_CHROME_VERSION = 102 # What chromedriver version to start at

# Browser to be used
BROWSER = 'Undetected Chrome'  # Options: 'Chrome', 'Undetected Chrome'

# Action booleans
use_chromedriver_autoinstaller = True  # Production value: True
download_leads_file = True  # Production value: True
upload_leads_to_bt = True  # Production value: True
reset_google_sheets = True  # Production value: True
pause_on_error = False  # Production value: False
send_status_email = True  # Production value: True
send_notification_email = True  # Production value: True


# Clear the terminal
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
        if use_chromedriver_autoinstaller:
            chromedriver_autoinstaller.install()  # Check if the current version of chromedriver exists
                                      # and if it doesn't exist, download it automatically,
                                      # then add chromedriver to path
            #Removes navigator.webdriver flag to allow fully undetected
            option = webdriver.ChromeOptions()
            option.add_argument('--disable-blink-features=AutomationControlled')
            #Open Browser
            if BROWSER == 'Chrome':
                driver = webdriver.Chrome(options=option)
            elif BROWSER == 'Undetected Chrome':
                driver = uc.Chrome(options=option)
        else:
            # Download and install latest chromedriver
            installLatestChromedriver()
            # Try identified version
            set_chrome_version_global()
            directory = get_config_data(DataCategories.CHROMEDRIVER_DATA, 'directory').replace('/', '\\')
            chromedriver_path = (os.path.realpath(__file__)[::-1][(os.path.realpath(__file__)[::-1].find('\\')+1):])[::-1] + directory + 'chromedriver_' + str(chrome_version) + '_win.exe'
            try:
                if BROWSER == 'Chrome':
                    driver = webdriver.Chrome(executable_path=chromedriver_path)
                elif BROWSER == 'Undetected Chrome':
                    driver = uc.Chrome(executable_path=chromedriver_path)
            except WebDriverException:
                # Try latest chromedriver
                latest = 0
                for filename in os.listdir('chromedrivers'):
                    if filename.startswith('chromedriver_'): 
                        if int(filename[13:16]) > latest: latest = int(filename[13:16])
                chromedriver_path = (os.path.realpath(__file__)[::-1][(os.path.realpath(__file__)[::-1].find('\\')+1):])[::-1] + directory + 'chromedriver_' + str(latest) + '_win.exe'
                try:
                    if BROWSER == 'Chrome':
                        driver = webdriver.Chrome(executable_path=chromedriver_path)
                    elif BROWSER == 'Undetected Chrome':
                        driver = uc.Chrome(executable_path=chromedriver_path)
                except WebDriverException:
                    raise ValueError('Could not open chrome with given chromedriver.\nChrome version: ' + str(latest) + '\nChromeDriver path: ' + chromedriver_path)
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
        raise ValueError('Could not copy leads download file to archive folder: ' + str(e))
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
    tries = 3
    success = False
    for i in range(tries):
        # Go to the login url
        driver.get(get_config_data(DataCategories.BT_LOGIN_DATA, 'login_url'))
        # Log in
        driver.find_element(By.XPATH, '//*[@id="username"]').send_keys(get_config_data(DataCategories.BT_LOGIN_DATA, 'username'))  # Enter username
        driver.find_element(By.XPATH, '//*[@id="password"]').send_keys(get_config_data(DataCategories.BT_LOGIN_DATA, 'password'))  # Enter password
        # Click login button
        login_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#reactLoginListDiv button')))
        actionChains.move_to_element(login_btn).click().perform()
        # Wait for page to load
        try:
            WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#reactMainNavigation')))
            success = True
            break
        except TimeoutException:
            pass
    if not success:
        raise ValueError('Could not login after ' + str(tries) + ' tries')


# Upload leads file to BuilderTrend
def btUploadLeads(leads_filepath):
    clear()
    print('Attempting leads upload')
    # Go to BT leads url
    driver.get(get_config_data(DataCategories.LEADS_DATA, 'leads_bt_url'))
    # Click import leads button
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="rc-tabs-0-panel-ListView"]/header/a[2]/button'))).click()
    # Select file to upload
    file_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.ImportWizard .UploadButton input')))
    file_input.send_keys(leads_filepath)
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
            err_msg = ''
            for elem in error_container.find_elements(By.CSS_SELECTOR, 'div.ant-alert-description > ul > li'):
                err_msg += elem.text + '\n' 
            # Check if it is an email error
            if "This contact's primary email is associated with another account" in err_msg:
                error_rows = []
                for err in err_msg.split('\n'):
                    match = re.search(r'Row\s\d+,', err)
                    try:
                        match2 = re.search(r'\d+', str(match.group()))
                    except AttributeError:
                        pass
                    else:
                        error_rows.append(int(match2.group()))
                moveEmailsToNotes(leads_filepath, error_rows)
                btUploadLeads(leads_filepath)
                return
            else:
                print('ERROR OCCURED:', err_msg)
                raise ValueError(err_msg)
        # Attempt to click "Next" button
        try:
            itr += 1
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.ImportWizard [data-testid="next"]'))).click()
        except (StaleElementReferenceException, ElementClickInterceptedException, TimeoutException): 
            try:
                # Check for import wizard results container
                driver.find_element(By.CSS_SELECTOR, '.ImportWizard .importWizardResults')
            except NoSuchElementException:
                pass
            else:
                break
        time.sleep(1)
    # Check if import was successful
    try:
        WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.ImportWizard .success-message')))
    except (NoSuchElementException, TimeoutException):
        err_msg = 'Upload was NOT successful. No element could be found matching ".ImportWizard .success-message"'
        print(err_msg)
        raise ValueError(err_msg)
    else:  # Import was successful
        print('Upload was successful!')
        if send_notification_email: sendEmail('BuilderTend Lead Importer: NOTIFICATION', '\nNew leads have been uploaded to BuilderTrend!\n\nhttps://buildertrend.net/Leads/LeadsList.aspx', [get_config_data(DataCategories.EMAIL_DATA, 'steve_email')])
        # Check for skipped records
        try:
            driver.find_element(By.XPATH, '//button/span[text()="Download skipped records"]').click()
        except NoSuchElementException:
            pass  # No skipped records
        else:
            # Read xls file (skipped records)
            downloads_dir = 'C:\\Users\\btauto\\Downloads\\'
            xls_filename = 'LeadsImportReport.xls'
            xls_filepath = downloads_dir + xls_filename
            read_file = pd.read_excel(xls_filepath)
            # Write the dataframe object into csv file
            csv_filename = 'LeadsImportReport.csv'
            csv_filepath = downloads_dir + csv_filename
            read_file.to_csv (csv_filepath, index = None, header=True)
            # Delete LeadsImportReport.xls
            os.remove(xls_filepath)
            # Raise error showing skipped leads
            skipped_lines = None
            with open(csv_filepath, 'r') as f:
                skipped_lines = f.readlines()
                skipped_leads_str = '\n'.join(skipped_lines)
                raise ValueError('Most leads were imported, however some leads were skipped:\n' + skipped_leads_str)


# Finds any duplicate emails in csv and then moves them to general notes column
def fixDuplicateEmails(filepath:str):
    # Read file
    with open(filepath, 'r') as f:
        lines = f.readlines()
        colnames = lines[0].split(',')
        email_col_index = colnames.index('Email')
        emails = []
        duplicate_email_rows = []
        for i in range(len(lines)):
            if lines[i][0] != ',':
                line = lines[i].split(',')
                if line[email_col_index] not in emails:
                    emails.append(line[email_col_index])
                else:
                    duplicate_email_rows.append(i-1)
    moveEmailsToNotes(filepath, duplicate_email_rows)


# Moves lead emails from the email column to the end of the General notes column for every row given (1st lead is on row 0)
def moveEmailsToNotes(filepath:str, rows:list):
    new_csv_str = ''
    # Read file
    with open(filepath, 'r') as f:
        lines = f.readlines()
        colnames = lines[0].split(',')
        email_col_indx = colnames.index('Email')
        notes_col_indx = colnames.index('General Notes')
        new_lines = []
        for i in range(len(lines)):
            if lines[i][0] != ',':
                if i-1 in rows:
                    line = lines[i].split(',')
                    line[notes_col_indx] += '  ' + line[email_col_indx]
                    line[email_col_indx] = ''
                    new_lines.append(','.join(line))
                else:
                    new_lines.append(lines[i])
        new_csv_str = ''.join(new_lines)

    # Write file
    with open(filepath, 'w') as outfile:
        outfile.write(new_csv_str)


# Login to EH gmail account
def googleLogin():
    try:
        driver.get('https://accounts.google.com/')
        # enter email and password
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#identifierId'))).send_keys(get_config_data(DataCategories.EMAIL_DATA, 'eh_email'))
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#identifierNext button'))).click()
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#password input'))).send_keys(get_config_data(DataCategories.EMAIL_DATA, 'eh_email_password'))
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
    if not pause_on_error: os.system('cls' if os.name == 'nt' else 'clear')
    print('Closing driver...')
    try:
        driver.quit()
    except AttributeError:
        # Driver not created
        pass
    print('\n\nDriver Closed\n')
    sys.exit()


# Runs the program with manual lead csv file choosing and no google sheets resetting
def manual():
    try:
        root = Tk()
        root.withdraw()
        filename = askopenfilename()
        root.destroy()
        if not filename:
            print('No file was selected')
            sys.exit()
        fixDuplicateEmails(filename)
        initDriver()
        initActionChains()
        btLogin()
        btUploadLeads(filename)
        closeChrome()
    except Exception as e: 
        if pause_on_error:
            print(str(e))
            input('Press [Enter]... ')
        closeChrome()


# Runs complete program 
def main():
    try:
        initDriver()
        initActionChains()
        if download_leads_file:
            downloadLeadsFile()
        if upload_leads_to_bt:
            fixDuplicateEmails(get_config_data(DataCategories.LEADS_DATA, 'leads_download_filepath'))
            btLogin()
            btUploadLeads(get_config_data(DataCategories.LEADS_DATA, 'leads_download_filepath'))
        if reset_google_sheets:
            googleLogin()
            clearSheets()
        if send_status_email: newStatus('Program ran successfully', False)
        closeChrome()
    except Exception as e:
        # Get leads
        leads_str = ''
        with open(get_config_data(DataCategories.LEADS_DATA, 'leads_download_filepath')) as f:
            lines = f.readlines()
            for line in lines:
                if line[0] != ',':
                    leads_str += line
        print('Error encountered')
        if send_status_email: newStatus('ERROR encountered while running program:\n ' + str(e) + '\n\nLeads that were not imported:\n' + leads_str, True)
        if pause_on_error:
            print(str(e))
            input('Press [Enter]... ')
        closeChrome()


# Main method
if __name__ == '__main__':
    args = sys.argv[1:]
    if not args:
        main()
    if len(args) == 1 and (args[0] == '-m' or args[0] == '--manual'):
        manual()
    elif len(args) == 1 and (args[0] == '-h' or args[0] == '--help'):
        print("""
            {-m|--manual} - runs the program with manual lead csv file choosing and no google sheets resetting\n
            {-h|--help} - displays up this help page\n
        """)
        sys.exit()
    else:
        print('Command line args were invalid')
        sys.exit()
