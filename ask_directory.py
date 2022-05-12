import os
from tkinter import Tk, Label, Button
from tkinter import filedialog as fd

window = None
directory_path = ''
curr_directory = ''

# Ask user to specify a folder
def askForDirectory(description: str, current_directory: str = None):
    # Ask user to choose a folder
    global window
    window = Tk()
    window.title('Select ' + description + ' folder')
    lbl = Label(window, text='Select ' + description + ' folder')
    lbl.pack()
    if current_directory:
        global curr_directory
        curr_directory = current_directory
        curr_dir_lbl = Label(window, text='Current Directory:  ' + current_directory)
        curr_dir_lbl.pack()
        curr_dir_lbl.focus()
        new_dir_btn = Button(window, text='Change', command=newDir)
        new_dir_btn.pack()
        window.bind('<Return>', keepDir)  # Bind Enter button to setInput()
        keep_dir_btn = Button(window, text='Keep', command=keepDir)
        keep_dir_btn.pack()
        keep_dir_btn.focus()
    else:
        new_dir_btn = Button(window, text='Select Folder', command=newDir)
        new_dir_btn.pack()
        new_dir_btn.focus()
        window.bind('<Return>', newDir)  # Bind Enter button to setInput()
    window.state('zoomed')  # Maximize tk window
    window.mainloop()
    return directory_path

# When select Folder btn was clicked
def newDir(*args):
    global directory_path
    directory_path = fd.askdirectory() # show an "Open" dialog box and return the path to the selected file
    if os.name == 'nt':  # Windows
        directory_path = directory_path.replace('/', '\\')
    window.destroy()

# Keep the currently selected directory
def keepDir(*args):
    global directory_path
    directory_path = curr_directory
    window.destroy()

# Main
if __name__ == '__main__':
    askForDirectory()