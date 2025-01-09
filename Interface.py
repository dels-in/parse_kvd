import tkinter as tk
from tkinter import filedialog

import ResearchExtraction


def create_interface():
    def open_path_dialog(module_function):
        folder_path = filedialog.askdirectory()
        module_function(folder_path)

    root = tk.Tk()
    root.title("Выбор модуля")
    # Создаем кнопки для выбора модуля
    button1 = tk.Button(root, text="Извлечь ГКИ", command=lambda: open_path_dialog(ResearchExtraction.process_files_in_directory))
    # Размещаем кнопки в окне
    button1.pack(pady=10)
    root.mainloop()


# Запуск интерфейса
create_interface()
