import zipfile, requests, os

def installLatestChromedriver():
    version = requests.get('https://chromedriver.storage.googleapis.com/LATEST_RELEASE').text
    version_1st_num = version[:version.find('.')]
    url = 'https://chromedriver.storage.googleapis.com/{0}/{1}'.format(version, 'chromedriver_win32.zip')
    r = requests.get(url, allow_redirects=True)
    open('chromedriver.zip', 'wb').write(r.content)
    with zipfile.ZipFile("chromedriver.zip", "r") as zip_ref:
        zip_ref.extractall()
    new_chromedriver_filepath = (os.path.realpath(__file__)[::-1][(os.path.realpath(__file__)[::-1].find('\\')+1):])[::-1] + 'chromedriver_' + version_1st_num + '_win.exe'
    if os.path.exists(new_chromedriver_filepath): os.remove(new_chromedriver_filepath)
    os.rename('chromedriver.exe', new_chromedriver_filepath)
    os.remove('chromedriver.zip')