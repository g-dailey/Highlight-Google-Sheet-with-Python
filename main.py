
qa_date = "2021-04-16" #@param {type:"date"}
qa_sheet_name = "April 12,14,16" #@param {type:"string"}
agent_list_sheet = "Sheet_nameame - Month "
qa_sheet_link = "sheet_link" #@param {type:"string"}
agent_list_sheet = "Sheet_name - Month" #@param {type:"string"}

# !pip install --upgrade gspread
# !pip install gspread-formatting
from datetime import datetime
import random
import pandas as pd
import numpy as np
from google.colab import auth
auth.authenticate_user()
import gspread
from gspread_formatting import *
from oauth2client.client import GoogleCredentials
gc = gspread.authorize(GoogleCredentials.get_application_default())


qa_wb = gc.open_by_url(qa_sheet_link)
random_agent_wb = gc.open_by_key('ENTER_YOUR_GOOGLE_SHEET_KEY')
random_agent_sheet = random_agent_wb.worksheet(agent_list_sheet)
random_qa_values = random_agent_sheet.get_all_values()

random_agent_df = pd.DataFrame.from_records(random_qa_values[1:], columns=random_qa_values[0])
date_random_qa_cell = random_agent_sheet.find('{dt.month}/{dt.day}/{dt.year}'.format(dt=datetime.strptime(qa_date, '%Y-%m-%d')))
qa_ws = qa_wb.worksheet(qa_sheet_name)
date_qa_cell = qa_ws.find(datetime.strptime(qa_date, '%Y-%m-%d').strftime('%A, %B %-d, %Y'))

start_row, end_row = 4, qa_ws.row_count
start_col = date_qa_cell.col

qa_lst = qa_ws.get(f"{gspread.utils.rowcol_to_a1(start_row, start_col)}:{gspread.utils.rowcol_to_a1(end_row, start_col+18)}")
qa_lst = [row[0].lower() for row in qa_lst[1:] if len(row) > 0]
random_agent_df['Auditor']= random_agent_df['Auditor'].str.strip().str.lower()
diff = set(random_agent_df['Auditor']) - set(qa_lst)

print('Below agents are not in the QA list for 'f'{qa_date}, please contact the vendor and request them to QA the agents for the mentioned date.')
for email in
  print(email)
print(len(diff))


qa_ws_values = qa_ws.get(f"{gspread.utils.rowcol_to_a1(start_row, start_col)}:{gspread.utils.rowcol_to_a1(end_row, start_col+1)}")

qa_df = pd.DataFrame(qa_ws_values[1:], columns=qa_ws_values[0])
qa_ws_values = [{ "row_number": i, "raw": row } for i,row in enumerate(qa_ws_values, start_row)]

grouped_lst = qa_df.groupby('Auditor').first()
qa_random_lst = grouped_lst.sample(frac = 0.25, replace=False)
qa_agent_lst = qa_random_lst.index.to_list()
qa_agent_df_final = pd.DataFrame(qa_agent_lst)

random_lst_export = random_agent_wb.worksheet(agent_list_sheet)
random_lst_export.update(gspread.utils.rowcol_to_a1(3,date_random_qa_cell.col),qa_agent_df_final.to_numpy().tolist())

qa_random_set = set(qa_random_lst.index)

fmt = cellFormat(
    backgroundColor=color(4, 4, 0),
    )

with batch_updater(qa_ws.spreadsheet) as batch:
  is_highlighted = set()
  for row in qa_ws_values:
    if len(row['raw']) == 0:
      continue
    auditor_name = row['raw'][0]
    if auditor_name in is_highlighted:
      continue
    should_highlight = auditor_name in qa_random_set
    if should_highlight:
      batch.format_cell_range(qa_ws,f"{gspread.utils.rowcol_to_a1(row['row_number'], start_col)}:{gspread.utils.rowcol_to_a1(row['row_number']+24, start_col)}", fmt)
      is_highlighted.add(auditor_name)
