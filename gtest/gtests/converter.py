import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials

def excel_to_google_sheet(email, excel_file_path):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('C:/Users/lev/Documents/study/UPPRPO/Project_OOP/creditionals_2.json', scope)
    gc = gspread.authorize(credentials)
    
    sh = gc.create('New Spreadsheet')
    sh.share(email, perm_type='user', role='writer')
    
    xl = pd.ExcelFile(excel_file_path)
    sheet_names = xl.sheet_names
    
    for sheet_name in sheet_names:
        excel_data = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        data = excel_data.values.tolist()
        worksheet = sh.add_worksheet(title=sheet_name, rows=len(data) + 1, cols=len(data[0]))
        worksheet.append_row(excel_data.columns.tolist())
        for row in data:
            worksheet.append_row(row)
    
    sh.del_worksheet(sh.sheet1)
    
    return sh.url

def google_sheet_to_excel(google_sheet_url, excel_file_path):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('C:/Users/lev/Documents/study/UPPRPO/Project_OOP/creditionals_2.json', scope)
    gc = gspread.authorize(credentials)
    
    sh = gc.open_by_url(google_sheet_url)
    worksheet_list = sh.worksheets()
    
    with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as excel_writer:
        for worksheet in worksheet_list:
            data = worksheet.get_all_values()
            if not data:
                continue
            df = pd.DataFrame(data[1:], columns=data[0])
            df.to_excel(excel_writer, sheet_name=worksheet.title, index=False)