import json, os
from pathlib import Path
from tkinter import Tk, Label, Text, Button
from ask_filepath import askForFilePath

class DataCategories():
    INIT_DATA = 'init_data'
    BT_LOGIN_DATA = 'bt_login_data'
    CHROMEDRIVER_DATA = 'chromedriver_data'
    STATUS_EMAIL_DATA = 'status_email_data'

# Data
init_data = {
    'data_name': 'init',
    'description': 'Initialization',
    'bot_path': 'src/bot/bot.py',  # relative to the run_bot.bat file
    'config_path': 'src/bot/config.py', # relative to the run_bot.bat file
    'leads_file_path': None
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
    'directory': '/'  # relative to the bot.py file
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
    dirname = Path(os.path.dirname(__file__))
    filepath = os.path.join(dirname, 'data\\' + category + '.json')
    data_dict = JSONtoDict(filepath)
    if data_dict is None:
        data_dict = writeJSON(globals()[category])
    for k in data_dict:
        if k == key:
            return data_dict[key]
    raise ValueError('the parameter \'key\' given (' + key + ') does not exist in ' + category)
    

# Write a data dict to a json file
def writeJSON(data_dict: dict):
    # Look through dict for null values
    for key in data_dict:
        if data_dict[key] is None:
            if not key.find('path'):
                window = Tk()
                window.title(data_dict['description'] + ' | ' + key)
                window.geometry('400x100')
                lbl = Label(window, text='Please enter ' + data_dict['description'] + ' ' + key + ': ')
                txt = Text(window, height = 1, width = 100)
                def setInput(e):
                    input = txt.get('1.0', 'end-2c')
                    data_dict[key] = input
                    window.destroy()
                window.bind('<Return>', setInput)  # Bind Enter button to setInput()
                btn = Button(window, text='Submit', command=lambda:setInput(''))
                lbl.pack()
                txt.pack()
                btn.pack()
                window.state('zoomed')  # Maximize tk window
                window.mainloop()
            else:
                askForFilePath('Leads Excel File')
    
    # Serializing jsons
    json_object = json.dumps(data_dict, indent = 4)
  
    # Writing to .json
    dirname = Path(os.path.dirname(__file__))
    filepath = os.path.join(dirname, 'data\\' + data_dict['data_name'] + '_data.json')
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
    # Update the init_data.json file
    writeJSON(init_data)
