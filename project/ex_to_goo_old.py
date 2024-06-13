import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import argparse


def from_excel_to_google(excel_file_path, google_sheet_url):
    # Авторизация в Google Sheets API
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('C:/Users/lev/Documents/study/UPPRPO/Project_OOP/creditionals_2.json', scope)
    gc = gspread.authorize(credentials)
    
    # Открываем Google Sheets по ссылке
    sh = gc.open_by_url(google_sheet_url)
    
    # Получаем первый лист в Google Sheets
    worksheet = sh.get_worksheet(0)  # 0 - индекс листа
    
    # Читаем данные из Excel файла с помощью pandas
    excel_data = pd.read_excel(excel_file_path)
    
    # Преобразуем данные в формат списка списков (двумерный массив)
    data = excel_data.values.tolist()
    
    # Очищаем старые данные в Google Sheets
    worksheet.clear()
    
    # Записываем новые данные в Google Sheets
    for row in data:
        worksheet.append_row(row)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Convert xls to Google Sheets')
    parser.add_argument('input_path', type=str, help='Path to the input XLS file')
    parser.add_argument('google_sheet_id', type=str, help='ID of the Google Sheet')
    parser.add_argument('email', type=str, help='Email address for sharing the Google Sheet')

    args = parser.parse_args()
    from_excel_to_google(args.input_path, args.google_sheet_id)