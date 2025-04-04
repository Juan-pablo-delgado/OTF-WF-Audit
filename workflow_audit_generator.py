from datetime import datetime
import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

# Abrir el diálogo para seleccionar un archivo
file_path = filedialog.askopenfilename(
    title="Selecciona un archivo de Excel",
    filetypes=[("Archivos de Excel", "*.xls;*.xlsx")]  # Filtro para archivos Excel
)

# Expresión regular para extraer el nombre del archivo
file_name = re.search(r"[^/\\]+(?=\.[a-zA-Z0-9]+$)", file_path).group()


account = input("Enter Account number: ")
bearer_token = input("Enter bearer token: ")

df = pd.read_excel(file_path)

#Agregar Workflow Name + link
def wf_apply_link(flow_id,wf_name,account):
    name = wf_name.replace('"',"'")
    return f'=HYPERLINK("https://app.hubspot.com/workflows/{account}/platform/flow/{flow_id}/edit","{name}")'

#Agregar Folder + link
def folder_apply_link(folder_id,account):
  if folder_id != 0:
    return f'=HYPERLINK("https://app.hubspot.com/workflows/{account}/folders?folderId={folder_id}","{folder_id}")'
  else:
    return '-'

#Agregar On/Off
def on_off(on_off):
  return 'On' if on_off else 'Off'

#Agregar index
df.reset_index(drop=True, inplace=True)
df.index = df.index + 1

#Issues?
def is_issue(issue):
  return 'Yes' if issue > 0 else 'No'

#Recommended Action
def recommended_action(is_issue, on_off, enrolled_last, last_action, total_enrolled):
  from datetime import timedelta
  is_older_than_one_year = datetime.now() - last_action > timedelta(days=365)

  if is_issue == 0 and on_off and enrolled_last > 0 and total_enrolled > 0:
      return 'Keep'
  elif enrolled_last > 0 or total_enrolled > 0 and on_off:
      return 'Review & Keep'
  elif not on_off and enrolled_last == 0 and total_enrolled == 0 and is_older_than_one_year:
      return 'Delete'
  else:
      return 'Review to Delete'

#Recommendations
def recommendation(recommended_action, issues):
    if recommended_action == 'Keep' and issues:
        return 'Everything is functioning properly. No issues have been detected, and registration activity is stable. You may proceed with regular operations.'
    elif recommended_action == 'Review & Keep' and issues:
        return 'An issue has been identified, but the system remains operational, and registrations are still functioning. It is advised to review the issue to ensure it does not affect performance or user experience. In the meantime, operations can continue until it is resolved.'
    elif recommended_action == 'Review to Delete':
        return 'The system appears inactive, with no registrations or activity. We recommend its removal if there are no plans for reactivation, as this will help maintain a clean and efficient database.'
    elif issues:
        return 'The system appears inactive, with no registrations or activity. We recommend deleting it if there are no plans to reactivate it, to help maintain a clean and efficient database.'

#Re-enrollment
def re_enrollment(flow_id):
  import requests

  url = f"https://api.hubapi.com/automation/v4/flows/{flow_id}"
  headers = {
      'accept': "application/json",
      'authorization': f'Bearer {bearer_token}'
     }
  try:
    response = requests.request("GET", url, headers=headers)
    if response.json().get('enrollmentCriteria').get('shouldReEnroll'):
      return 'Yes'
    else:
      return 'No'
  except:
    return ' '

df['Workflow Name'] = df.apply(lambda row: wf_apply_link(row['Flow ID'],row ['Name'],account), axis=1)
df['Folder Name'] = df.apply(lambda row: folder_apply_link(row['Folder'],account), axis=1)
df['ON/OFF'] = df.apply(lambda row: on_off(row['On or Off']), axis=1)
df['Issues?'] = df.apply(lambda row: is_issue(row['Current issues']), axis=1)
df['Recommended Action'] = df.apply(lambda row: recommended_action(row['Current issues'],row['On or Off'],row['Enrolled last 7-days'],row['Last action on'],row['Enrolled total']), axis=1)
df['Recommendation'] = df.apply(lambda row: recommendation(row['Recommended Action'],row['Issues?']), axis=1)
df['Re-enrollment'] = df.apply(lambda row: re_enrollment(row['Flow ID']), axis=1)
df['Month'] = pd.to_datetime(df['Created on']).dt.strftime('%B %Y')
df['Issue type'] = '-'
df['Issue details'] = '-'

# Crear las columnas del Audit

desired_columns = ['Workflow Name','Folder Name','ON/OFF','Created by', 'Created on','Month',
                   'Object type', 'Trigger Type', 'Enrolled total',
                   'Enrolled last 7-days','Last action on',
                   'Re-enrollment','Description','Issue type','Issue details', 'Issues?',
                   'Recommendation', 'Recommended Action']

audit_df = df[desired_columns]

try:
  audit_df.to_excel('Template HubSpot Audit.xlsx')
  print("Saved sussefully")
except:
  print("failed to generate")
  