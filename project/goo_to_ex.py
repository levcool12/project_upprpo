import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import argparse


def from_google_to_excel(google_sheet_url, excel_file_path):
       # Авторизация в Google Sheets API
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('C:/Users/lev/Documents/study/UPPRPO/Project_OOP/creditionals_2.json', scope)
    gc = gspread.authorize(credentials)
    
    # Открываем Google Sheets по ссылке
    sh = gc.open_by_url(google_sheet_url)
    
    # Получаем список всех листов в Google Sheets
    worksheet_list = sh.worksheets()
    
    # Создаем объект Pandas Excel writer с указанием пути к файлу Excel
    with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as excel_writer:
        for worksheet in worksheet_list:
            # Читаем данные из текущего листа Google Sheets
            data = worksheet.get_all_values()
            
            # Пропускаем пустые листы
            if not data:
                continue
            
            # Преобразуем данные в DataFrame
            df = pd.DataFrame(data[1:], columns=data[0])  # Первая строка в данных - заголовки
            
            # Записываем данные в Excel
            df.to_excel(excel_writer, sheet_name=worksheet.title, index=False)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Convert Google Sheets to xls')
    parser.add_argument('google_sheet_id', type=str, help='ID of the Google Sheet')
    parser.add_argument('output_path', type=str, help='Path to save the output XLS file')
    parser.add_argument('email', type=str, help='Email address for sharing the Google Sheet (not used in this script)')

    args = parser.parse_args()
    from_google_to_excel(args.output_path, args.google_sheet_id)