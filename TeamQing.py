import pandas as pd
from tkinter import *
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pprint

#Authorize the API
scope = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/drive.file'
    ]
file_name = '/Users/kevinjoseph/Downloads/PythonSheet-key.json'
creds = ServiceAccountCredentials.from_json_keyfile_name(file_name,scope)
client = gspread.authorize(creds)

sheet = client.open('PythonSheet Check-In').sheet1
df = pd.read_excel('/Users/kevinjoseph/Downloads/TeamsList.xls')

#Fill sheet with team names and check-in status
row = 2
sheet.update('A1', 'Team')
sheet.update('B1', 'Checked-In?')
for i in df.loc[:,'Team']:
    sheet.update_cell(row, 1, i)
    sheet.update_cell(row, 2, 'N')
    row += 1

python_sheet = sheet.get_all_records()
pp = pprint.PrettyPrinter()
pp.pprint(python_sheet)

class App(Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.pack()
        
# Create object
root = App()

# Adjust size
root.master.title("Team Check-In")
root.master.maxsize(400, 800)
root.master.geometry('300x400')

# Change the label text
def show():
    
    # Hide everything
    l1.pack_forget()
    drop.pack_forget()
    button.pack_forget()

    # Send confirmation and update sheets
    label.config( text = 'Submitted' )
    cell = sheet.find(clicked.get())
    sheet.update_cell(cell.row, cell.col + 1, 'Y')

    # Update dropdown options
    options.remove(clicked.get())
    menu = drop['menu']
    menu.delete(0,'end')
    for i in options:
        menu.add_command(label=i, command=lambda value=i: clicked.set(i))


    # Update clicked and redo watching
    clicked.set( "Select Team Number" )
    clicked.trace('w', watcher)
    
    # Return everything after 5 seconds
    label.after(1000, resetLabel)
    l1.after(1000, l1.pack)
    drop.after(1000,drop.pack)
    button.after(1000, resetButton)
    
# Func to reset the label
def resetLabel():
    label.pack_forget()
    label.config(text = '')
    label.after(10,label.pack)

# Func to reset the button
def resetButton():
    button.config(state = 'disabled')
    button.pack()
# Disable button till drop down is selected
def watcher(*args):
    button.config(state = 'normal')

# Dropdown menu options
options = df.loc[:,"Team"]
options = options.tolist()
# Create Label 1
l1 = Label(root, text = 'Select the team you wish to check-in')
l1.pack()

# datatype of menu text
clicked = StringVar(root)
clicked.set( "Select Team Number" )
clicked.trace('w', watcher)

# Create Dropdown menu
drop = OptionMenu( root , clicked , *options )
drop.pack()

# Create button
button = Button( root , text = "Submit" , command = show, state = 'disabled')
button.pack()

# Create Label 2
label = Label( root , text = " " )
label.pack()

# Execute tkinter
root.mainloop()
