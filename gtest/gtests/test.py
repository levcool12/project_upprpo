import pytest
from unittest.mock import patch
from converter import excel_to_google_sheet, google_sheet_to_excel

@patch('converter.gspread.authorize')
@patch('converter.pd.ExcelFile')
@patch('converter.pd.read_excel')
def test_excel_to_google_sheet(mock_read_excel, mock_excel_file, mock_authorize):
    # Моки для авторизации
    mock_gc = mock_authorize.return_value
    mock_sh = mock_gc.create.return_value
    
    # Моки для чтения Excel файла
    mock_excel_file.return_value.sheet_names = ['Sheet1']
    mock_read_excel.return_value.values.tolist.return_value = [[1, 2], [3, 4]]
    mock_read_excel.return_value.columns.tolist.return_value = ['A', 'B']
    
    email = 'dimakil20030310@gmail.com'
    excel_file_path = 'C:/Users/lev/Downloads/22_7.xlsx'
    
    result_url = excel_to_google_sheet(email, excel_file_path)
    
    # Проверка вызовов
    mock_authorize.assert_called_once()
    mock_gc.create.assert_called_once_with('New Spreadsheet')
    mock_sh.add_worksheet.assert_called_once_with(title='Sheet1', rows=3, cols=2)
    mock_sh.del_worksheet.assert_called_once()
    
    # Проверка результата
    assert isinstance(result_url, str)

@patch('converter.gspread.authorize')
@patch('converter.pd.ExcelWriter')
def test_google_sheet_to_excel(mock_excel_writer, mock_authorize):
    # Моки для авторизации
    mock_gc = mock_authorize.return_value
    mock_sh = mock_gc.open_by_url.return_value
    mock_worksheet = mock_sh.worksheets.return_value
    
    # Моки для работы с листами
    mock_worksheet[0].get_all_values.return_value = [['A', 'B'], [1, 2], [3, 4]]
    
    google_sheet_url = 'https://docs.google.com/spreadsheets/d/1wC9kI3J3IoK6zwWCeQnW1Pxi2WhKGUFnHdx-LdCQy5U/edit?gid=0#gid=0'
    excel_file_path = 'save.xlsx'
    
    google_sheet_to_excel(google_sheet_url, excel_file_path)
    
    # Проверка вызовов
    mock_authorize.assert_called_once()
    mock_gc.open_by_url.assert_called_once_with(google_sheet_url)
    mock_excel_writer.assert_called_once_with(excel_file_path, engine='xlsxwriter')
    
    # Проверка записи данных в Excel
    with mock_excel_writer.return_value as mock_writer:
        mock_writer.__enter__().return_value = mock_writer
        mock_writer.save.assert_called_once()