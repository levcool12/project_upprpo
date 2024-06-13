import gspread
from oauth2client.service_account import ServiceAccountCredentials
import chardet
import json

# Определение кодировки файла
def detect_encoding(file_path):
    with open(file_path, 'rb') as f:
        raw_data = f.read()
        result = chardet.detect(raw_data)
        return result['encoding']

# Укажите путь к вашему JSON файлу с учетными данными
json_file = 'C:/Users/admin/Desktop/Project_OOP/credentials.json'

# Определяем кодировку
encoding = detect_encoding(json_file)
print(f"Определенная кодировка: {encoding}")

# Чтение файла с ключом
try:
    with open(json_file, 'r', encoding=encoding) as f:
        creds_data = json.load(f)
except Exception as e:
    print(f"Ошибка при чтении JSON файла: {e}")
    exit(1)

# Настроим доступ к Google Sheets
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive"
]

try:
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_data, scope)
    client = gspread.authorize(creds)
except Exception as e:
    print(f"Ошибка при авторизации: {e}")
    exit(1)

# Создание новой Google Таблицы
try:
    spreadsheet = client.create('проба')
    spreadsheet.share('dimakil20030310@gmail.com', perm_type='user', role='writer')
    print(f"Создана новая таблица: {spreadsheet.url}")
except Exception as e:
    print(f"Ошибка при создании таблицы: {e}")
    exit(1)

# Дополнительное предоставление доступа к таблице по указанной почте
spreadsheetId = spreadsheet.id  # или spreadsheet['spreadsheetId'] в старой версии API

try:
    driveService = gspread.authorize(creds)
    driveService.insert_permission(
        spreadsheetId,
        'my_test_address@gmail.com',
        perm_type='user',
        role='writer'
    )
    print(f"Предоставлен доступ для my_test_address@gmail.com")
except Exception as e:
    print(f"Ошибка при предоставлении доступа: {e}")
    exit(1)
