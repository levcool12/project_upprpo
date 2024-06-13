import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import argparse

def excel_to_google_sheet(email, excel_file_path):
    # Авторизация в Google Sheets API
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('C:/Users/lev/Documents/study/UPPRPO/Project_OOP/creditionals_2.json', scope)
    gc = gspread.authorize(credentials)
    
    # Создаем новую Google таблицу
    sh = gc.create('New Spreadsheet')
    
    # Предоставляем доступ по email
    sh.share(email, perm_type='user', role='writer')
    
    # Получаем список всех листов в Excel файле
    xl = pd.ExcelFile(excel_file_path)
    sheet_names = xl.sheet_names
    
    for sheet_name in sheet_names:
        # Читаем данные из текущего листа Excel файла
        excel_data = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        
        # Преобразуем данные в формат списка списков (двумерный массив)
        data = excel_data.values.tolist()
        
        # Создаем новый лист в Google Sheets с именем текущего листа из Excel
        worksheet = sh.add_worksheet(title=sheet_name, rows=len(data) + 1, cols=len(data[0]))
        
        # Добавляем заголовки
        worksheet.append_row(excel_data.columns.tolist())
        
        # Записываем данные в новый лист Google Sheets
        for row in data:
            worksheet.append_row(row)
    
    # Удаляем первый пустой лист, созданный по умолчанию
    sh.del_worksheet(sh.sheet1)
    
    # Возвращаем ссылку на созданную таблицу
    return sh.url

# Пример использования
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Convert xls to Google Sheets')
    parser.add_argument('input_path', type=str, help='Path to the input XLS file')
    parser.add_argument('google_sheet_id', type=str, help='ID of the Google Sheet')
    parser.add_argument('email', type=str, help='Email address for sharing the Google Sheet')

    args = parser.parse_args()
    
    google_sheet_url = excel_to_google_sheet(args.email, args.input_path)
    print("Google Sheet created: ", google_sheet_url)