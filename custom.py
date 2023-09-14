from pathlib import Path
import PySimpleGUI as sg
import pandas as pd

# Add some color to the window
sg.theme('LightBrown13')

current_dir = Path(__file__).parent if '__file__' in locals() else Path.cwd()
EXCEL_FILE = current_dir / 'Data_Entry.xlsx'

# Load the data if the file exists, if not, create a new DataFrame
if EXCEL_FILE.exists():
    df = pd.read_excel(EXCEL_FILE)
else:
    df = pd.DataFrame()

# Define the font for text elements
font = ('Helvetica', 16)  # Change 'Helvetica' to your desired font family and 16 to your desired font size

layout = [
    [sg.Text('Please fill out the following fields:', font=font)],
    [sg.Text('Name', size=(15, 1), font=font), sg.InputText(key='Name')],
    [sg.Text('City', size=(15, 1), font=font), sg.InputText(key='City')],
    [sg.Text('Favorite Colour', size=(15, 1), font=font), sg.Combo(['Green', 'Blue', 'Red'], key='Favorite Colour')],
    [sg.Text('I speak', size=(15, 1), font=font),
     sg.Checkbox('Mandarin', key='Mandarin', font=font),
     sg.Checkbox('Malay', key='Malay', font=font),
     sg.Checkbox('English', key='English', font=font)],
    [sg.Text('No. of Children', size=(15, 1), font=font), sg.Spin([i for i in range(0, 16)],
                                                                    initial_value=0, key='Children')],
    [sg.Submit(font=font), sg.Button('Clear', font=font)]
]

# Replace 'icon.ico' with the path to your custom icon
window = sg.Window('Form - Ayie', layout, icon='icon.ico')

def clear_input():
    for key in values:
        window[key]('')
    return None

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        new_record = pd.DataFrame(values, index=[0])
        df = pd.concat([df, new_record], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)  # This will create the file if it doesn't exist
        sg.popup('Data disimpan!')
        clear_input()
window.close()
