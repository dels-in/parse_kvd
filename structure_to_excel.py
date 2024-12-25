import os
import tkinter as tk
from tkinter import filedialog
import pandas as pd


def create_excel(folder_path):
    def build_table(root, level=0):
        table = []
        for item in os.listdir(root):
            item_path = os.path.join(root, item)
            if os.path.isdir(item_path):
                row = [''] * level + [item]
                table.append(row)
                table.extend(build_table(item_path, level + 1))
            else:
                row = [''] * level + [item]
                table.append(row)
        return table

    table = build_table(folder_path)
    max_cols = max(len(row) for row in table)
    levels = ['Уровень вложенности']
    for i in range(2, max_cols + 1):
        levels.append(f'Уровень {i}')
    df = pd.DataFrame(table, columns=levels)
    df.to_excel('folder_structure.xlsx', index=False)


def browse_button():
    global folder_path
    folder_path = filedialog.askdirectory()
    create_excel(folder_path)


# Создание графического интерфейса
root = tk.Tk()
root.title("Выбор папки")
root.geometry("300x100")

browse_button = tk.Button(root, text="Выбрать папку", command=browse_button)
browse_button.pack(pady=20)

root.mainloop()
