import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess

class FileConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("File Converter")

        # Форматы файлов
        self.conversion_options = ["Excel to Google", "Google to Excel"]

        # UI элементы
        self.create_widgets()

    def create_widgets(self):
        # Выбор типа конвертации
        self.label_conversion_type = ttk.Label(self.root, text="Тип конвертации:")
        self.label_conversion_type.grid(column=0, row=0, padx=10, pady=10, sticky=tk.W)

        self.conversion_type = ttk.Combobox(self.root, values=self.conversion_options)
        self.conversion_type.grid(column=1, row=0, padx=10, pady=10, sticky=tk.W)
        self.conversion_type.current(0)

        # Путь к Excel файлу
        self.label_excel_path = ttk.Label(self.root, text="Путь к Excel файлу:")
        self.label_excel_path.grid(column=0, row=1, padx=10, pady=10, sticky=tk.W)

        self.excel_path = ttk.Entry(self.root, width=50)
        self.excel_path.grid(column=1, row=1, padx=10, pady=10, sticky=tk.W)

        self.button_excel_browse = ttk.Button(self.root, text="Обзор", command=self.browse_excel_file)
        self.button_excel_browse.grid(column=2, row=1, padx=10, pady=10, sticky=tk.W)

        # ID Google таблицы
        self.label_google_sheet_id = ttk.Label(self.root, text="ID Google таблицы:")
        self.label_google_sheet_id.grid(column=0, row=2, padx=10, pady=10, sticky=tk.W)

        self.google_sheet_id = ttk.Entry(self.root, width=50)
        self.google_sheet_id.grid(column=1, row=2, padx=10, pady=10, sticky=tk.W)

        # Почта пользователя
        self.label_email = ttk.Label(self.root, text="Почта пользователя:")
        self.label_email.grid(column=0, row=3, padx=10, pady=10, sticky=tk.W)

        self.email = ttk.Entry(self.root, width=50)
        self.email.grid(column=1, row=3, padx=10, pady=10, sticky=tk.W)

        # Кнопка конвертации
        self.button_convert = ttk.Button(self.root, text="Конвертировать", command=self.convert_file)
        self.button_convert.grid(column=1, row=4, padx=10, pady=20, sticky=tk.W)

    def browse_excel_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls *.xlsx")])
        self.excel_path.delete(0, tk.END)
        self.excel_path.insert(0, file_path)

    def convert_file(self):
        conversion_type = self.conversion_type.get()
        excel_path = self.excel_path.get()
        google_sheet_id = self.google_sheet_id.get()
        email = self.email.get()

        if not excel_path or not google_sheet_id or not email:
            messagebox.showerror("Ошибка", "Пожалуйста, заполните все поля.")
            return

        # Проверяем, какой конвертер нужно использовать
        if conversion_type == "Excel to Google":
            script = 'ex_to_goo.py'
        elif conversion_type == "Google to Excel":
            script = 'goo_to_ex.py'
        else:
            messagebox.showerror("Ошибка", "Неподдерживаемая конвертация.")
            return

        # Запускаем соответствующий скрипт с аргументами
        try:
            subprocess.run(['python', script, excel_path, google_sheet_id, email], check=True)
            messagebox.showinfo("Информация", f"Конвертация '{conversion_type}' выполнена успешно!")
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка при конвертации: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = FileConverterApp(root)
    root.mainloop()