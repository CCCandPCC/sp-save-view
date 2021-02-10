import download
import inquirer
from inquirer.themes import GreenPassion
import openpyxl
import os


print('''
   ____                _          _     _                 _     _          
  / ___|__ _ _ __ ___ | |__  _ __(_) __| | __ _  ___  ___| |__ (_)_ __ ___ 
 | |   / _` | '_ ` _ \| '_ \| '__| |/ _` |/ _` |/ _ \/ __| '_ \| | '__/ _ \\
 | |__| (_| | | | | | | |_) | |  | | (_| | (_| |  __/\__ \ | | | | | |  __/
  \____\__,_|_| |_| |_|_.__/|_|  |_|\__,_|\__, |\___||___/_| |_|_|_|  \___|
 |  _ \(_) __ _(_) |_ __ _| |             |___/                            
 | | | | |/ _` | | __/ _` | |                                              
 | |_| | | (_| | | || (_| | |                                              
 |____/|_|\__, |_|\__\__,_|_|                                              
          |___/                                                            

Cambridgeshire Digital Sharepoint Downloader 0.2
This tool can download the Excel exports from Sharepoint Views.
https://github.com/CCCandPCC/sharepoint-view-save''')

xlsx = [f for f in os.listdir('.') if os.path.isfile(f) and f.endswith('.xlsx')]

questions = [
  inquirer.List('source_file',
    message="Select the Excel file containing the Sharepoint export",
    choices=xlsx)
]

setup = inquirer.prompt(questions, theme=GreenPassion())
ws = download.DownloadSharepoint()
sheets = ws.open_xl(setup['source_file'])

if len(sheets) > 1:
  questions2 = [inquirer.List('worksheet', "Select the worksheet containing the data", sheets)]
  ws_name = inquirer.prompt(questions2, theme=GreenPassion())['worksheet']
else:
  ws_name = sheets[0]

ws.select_ws(ws_name)

questions = [
  inquirer.Checkbox(
    'folders', 
    "Select the headers you'd like to use as folders (use space to select, and choose them in the order you want them saved)", 
    list(map(lambda x: x[0], ws.headers))),
  inquirer.Text(
    'output',
    message="Which folder would you like to output downloaded files to",
    default="output"
  )
]

options = inquirer.prompt(questions, theme=GreenPassion())
ws.download_sharepoint_xl(options['output'], options['folders'])

input('\nComplete. Press Enter to exit.')