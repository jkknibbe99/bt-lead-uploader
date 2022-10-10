import json, os
from pathlib import Path
from tkinter import Tk, Label, Text, Button
from ask_filepath import askForFilePath
from ask_directory import askForDirectory

class DataCategories():
    INIT_DATA = 'init_data'
    BT_LOGIN_DATA = 'bt_login_data'
    CHROMEDRIVER_DATA = 'chromedriver_data'
    FIREFOXDRIVER_DATA = 'firefoxdriver_data'
    LEADS_DATA = 'leads_data'
    EH_GMAIL_LOGIN_DATA = 'eh_gmail_login_data'
    STATUS_EMAIL_DATA = 'status_email_data'

# Data
init_data = {
    'data_name': 'init',
    'description': 'Initialization',
    'bot_path': 'src/bot/bot.py',  # relative to the run_bot.bat file
    'config_path': 'src/bot/config.py', # relative to the run_bot.bat file
}

bt_login_data = {
    'data_name': 'bt_login',
    'description': 'BuilderTrend Login',
    'login_url': 'https://buildertrend.net/Leads/LeadsList.aspx',
    'username': None,
    'password': None
}

chromedriver_data = {
    'data_name': 'chromedriver',
    'description': 'Chromedriver',
    'version': 100,
    'directory': '/chromedrivers/'  # relative to the bot.py file
}

firefoxdriver_data = {
    'data_name': 'firefoxdriver',
    'description': 'Firefox Driver',
    'version': 0.31,
    'directory': '',  # relative to the bot.py file (no slash at start)
    'profile_path': '/Users/joshua/AppData/Roaming/Mozilla/Firefox/Profiles/32pouvxh.default'
}

leads_data = {
    'data_name': 'leads',
    'description': 'Leads Data',
    'leads_download_filepath': None,
    'leads_file_download_url': None,
    'leads_archive_directory': None,
    'new_lead_sheet_1_url': None,
    'new_lead_sheet_2_url': None,
    'bt_lead_upload_sheet_url': None
}

eh_gmail_login_data = {
    'data_name': 'eh_gmail_login',
    'description': 'EH Gmail Login Data',
    'email': None,
    'password': None
}

status_email_data = {
    'data_name': 'status_email',
    'description': 'Status Email',
    'sender': 'jknibbe.dev@gmail.com',
    'receiver': 'jknibbe.dev@gmail.com',
    'password': None,
}


# Get configuration data
def get_config_data(category: str, key: str):
    # print('here')
    dirname = Path(os.path.dirname(__file__))
    filepath = os.path.join(dirname, 'data\\' + category + '.json')
    data_dict = JSONtoDict(filepath)
    if data_dict is None:
        data_dict = writeJSON(globals()[category])
    for k in data_dict:
        if k == key:
            if data_dict[key] == None:
                return writeJSON(data_dict)[key]
            else:
                return data_dict[key]
    raise ValueError('the parameter \'key\' given (' + key + ') does not exist in ' + category)
    

# Write a data dict to a json file
def writeJSON(data_dict: dict):
    # Look through dict for null values
    for key in data_dict:
        if data_dict[key] is None:
            if key.find('filepath') > -1:  # If key contains 'filepath'
                data_dict[key] = askForFilePath(key)
            elif key.find('directory') > -1:  # If key contains 'directory'
                data_dict[key] = askForDirectory(key)
            else:
                window = Tk()
                window.title(data_dict['description'] + ' | ' + key)
                window.geometry('400x100')
                lbl = Label(window, text='Please enter ' + data_dict['description'] + ' ' + key + ': ')
                txt = Text(window, height = 1, width = 150)
                def setInput(e):
                    inpt = txt.get('1.0', 'end-1c')
                    if inpt.endswith('\n'):
                        inpt = inpt[:-1]
                    data_dict[key] = inpt
                    window.destroy()
                window.bind('<Return>', setInput)  # Bind Enter button to setInput()
                btn = Button(window, text='Submit', command=lambda:setInput(''))
                lbl.pack()
                txt.pack()
                btn.pack()
                window.state('zoomed')  # Maximize tk window
                txt.focus()
                window.mainloop()
    
    # Serializing jsons
    json_object = json.dumps(data_dict, indent = 4)
  
    # Check for data directory
    dirname = Path(os.path.dirname(__file__))
    data_dir_path = os.path.join(dirname, 'data')
    if not os.path.isdir(data_dir_path):  # If data directory does not exist, create it
        os.mkdir(data_dir_path)
    # Writing to .json
    filepath = data_dir_path + '\\' + data_dict['data_name'] + '_data.json'
    with open(filepath, "w") as outfile:
        outfile.write(json_object)
    return data_dict


# Get a dict from json file
def JSONtoDict(filepath: str):
    try:
        with open(filepath) as json_file:
            data_from_json = json.load(json_file)
            return data_from_json
    except FileNotFoundError:
        return None


if __name__ == "__main__":
    # Update the the chose data file
    categories = ['init_data', 'bt_login_data', 'chromedriver_data', 'firefoxdriver_data', 'leads_data', 'eh_gmail_login_data', 'status_email_data']
    query_str = 'Select a data category to update:\n'
    choice = ''
    for i in range(len(categories)):
        query_str += '\t(' + str(i) + ') ' + categories[i] + '\n'
    query_str += '\n'
    os.system('cls' if os.name == 'nt' else 'clear')
    while True:
        response = input(query_str)
        try:
            int(response)
        except ValueError:
            if response in categories:
                writeJSON(globals()[response])
                choice = response
                break
            else:
                os.system('cls' if os.name == 'nt' else 'clear')
                print('Invalid selection. Try again.')
        else:
            response = int(response)
            if response < len(categories):
                writeJSON(globals()[categories[response]])
                choice = categories[response]
                break
            else:
                os.system('cls' if os.name == 'nt' else 'clear')
                print('Invalid selection. Try again.')
    os.system('cls' if os.name == 'nt' else 'clear')
    print(choice, 'updated successfully')
    