import os, sys
from tkinter import Tk
from tkinter.filedialog import askdirectory

root = Tk()
root.withdraw()
archive_dir = askdirectory()
root.destroy()
if not archive_dir:
    print('No folder was selected')
    sys.exit()

init = True

new_file_string = ''

for filename in os.listdir(archive_dir):
    filepath = archive_dir + '\\' + filename
    with open(filepath, 'r') as f:
        lines = f.readlines()
        for i in range(len(lines)):
            if i == 0:
                if init:
                    new_file_string += lines[0]
                    init = False
            elif lines[i][0] != ',':
                new_file_string += lines[i]

with open(archive_dir + '\\joined.csv', 'w') as f:
    f.write(new_file_string)

print('Program Complete')