import download
from PyInquirer import prompt
import openpyxl
import os
from requests_ntlm import HttpNtlmAuth

def main():
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

  Cambridgeshire Digital Sharepoint Downloader 0.4
  This tool can download the Excel exports from Sharepoint Views.

  https://github.com/CCCandPCC/sharepoint-view-save
  ''')

  xlsx = [f for f in os.listdir('.') if os.path.isfile(f) and f.endswith('.xlsx')]

  questions = [
    {
      'type':'list',
      'name': 'source_file',
      'message':"Select the Excel file containing the Sharepoint export",
      'choices': xlsx
    }
  ]

  if (len(xlsx) == 0):
    input('\nERROR: You must place the .xlsx file you want to import in the same directory as this application.')
    exit()

  setup = prompt(questions)
  ws = download.DownloadSharepoint()
  sheets = ws.open_xl(setup['source_file'])

  if len(sheets) > 1:
    questions2 = [{'type':'list', 'name':'worksheet', 'message':"Select the worksheet containing the data", 'choices':sheets}]
    ws_name = prompt(questions2)['worksheet']
  else:
    ws_name = sheets[0]

  ws.select_ws(ws_name)

  # authentication
  ws.do_auth = login

  # headers
  selected_folders = prompt([
    {
      'type': 'checkbox',
      'name': 'folders', 
      'message': "Select the headers you'd like to use as folders", 
      'choices': map(lambda x: {'name': str(x[0])}, ws.headers)
    }])['folders']

  ordered_folders = folder_order(selected_folders, True)

  out = prompt([{
      'type': 'input',
      'name': 'output',
      'message': "Which folder would you like to output downloaded files to",
      'default':"output"
    }
  ])['output']

  path_str = [out] + list(map(lambda x: '[' + x + ']', ordered_folders))
  msg = "Save files into " + '\\'.join(path_str) + "?"
  confirm = [{
    'type': 'confirm',
    'name': 'conf',
    'message': msg
  }]

  if prompt(confirm)['conf']:
    ws.download_sharepoint_xl(out, ordered_folders)

  input('\nComplete. Press Enter to exit.')

def login(session):
  print("\nYour login details aren't recognised. You must authenticate with Sharepoint to continue.\n")
  deets = prompt([
    {'type':'input', 'name':'user', 'message':'Enter the username to login to Sharepoint'},
    {'type': 'password', 'name': 'pass', 'message': 'Enter the password for the account'}
  ])
  session.auth = HttpNtlmAuth(deets['user'], deets['pass'])

def folder_order(choices, first):
  if len(choices) <= 1:
    return choices

  if first:
    msg = "Which header would you like to use as the top-level folder?"
  else:
    msg = "Which header should be used for the next level down of folder?"

  choice = prompt([{
    'type': 'list',
    'name': 'order',
    'message': msg,
    'choices': choices
  }])['order']

  choices.remove(choice)

  return [choice] + folder_order(choices, False)

try:
  main()
except Exception as e:
  print(f"\nERROR: {repr(e)}")
  input("\nPress Enter to exit.")