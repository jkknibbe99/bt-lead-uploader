import os
from tkinter import Tk, Label, Button
from tkinter import filedialog as fd

window = None
directory_path = ''
curr_directory = ''

# Ask user to specify a file
def askForFilePath(description: str, current_directory: str = None):
    # Ask user to choose a file
    global window
    window = Tk()
    window.title('Select ' + description)
    lbl = Label(window, text='Select ' + description)
    lbl.pack()
    if current_directory:
        global curr_directory
        curr_directory = current_directory
        curr_dir_lbl = Label(window, text='Current Directory:  ' + current_directory)
        curr_dir_lbl.pack()
        curr_dir_lbl.focus()
        new_dir_btn = Button(window, text='Change', command=newPath)
        new_dir_btn.pack()
        window.bind('<Return>', keepPath)  # Bind Enter button to setInput()
        keep_dir_btn = Button(window, text='Keep', command=keepPath)
        keep_dir_btn.pack()
        keep_dir_btn.focus()
    else:
        new_dir_btn = Button(window, text='Select File', command=newPath)
        new_dir_btn.pack()
        new_dir_btn.focus()
        window.bind('<Return>', newPath)  # Bind Enter button to setInput()
    window.state('zoomed')  # Maximize tk window
    window.mainloop()
    return directory_path

# When select File btn was clicked
def newPath(*args):
    global directory_path
    directory_path = fd.askopenfilename() # show an "Open" dialog box and return the path to the selected file
    if os.name == 'nt':  # Windows
        directory_path = directory_path.replace('/', '\\')
    window.destroy()

# Keep the currently selected directory
def keepPath(*args):
    global directory_path
    directory_path = curr_directory
    window.destroy()

# Main
if __name__ == '__main__':
    askForFilePath()