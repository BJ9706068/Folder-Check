import os
import tkinter as tk
from tkinter import ttk
import pandas as pd

def calculate_file_size(file_path):
    file_size = os.path.getsize(file_path)
    file_size_kb = round(file_size / 1024, 2)
    return file_size_kb

def generate_catalogue():
    folder_path = folder_path_var.get()

    # 显示程序运行结果
    result_label.config(text="Program started running.")

    # 遍历文件夹及其所有子文件夹，并检查是否为空
    folder_info = []
    for root, dirs, files in os.walk(folder_path):
        if len(files) == 0:
            folder_info.append([root, os.path.basename(root), "！！！Folder Is Empty！！！", ""])
        elif len(files) >= 1:
            folder_name = os.path.basename(root)
            for file in files:
                file_name = os.path.join(root, file)
                file_size_kb = calculate_file_size(file_name)
                folder_info.append([root, folder_name, file, file_size_kb])

    # 显示程序运行结果
    result_label.config(text="Program finished running.")

    # 生成目录列表Excel文件
    output_excel_file_path = os.path.join(folder_path, "Catalogue List.xlsx")
    with pd.ExcelWriter(output_excel_file_path, engine='xlsxwriter') as writer:
        df = pd.DataFrame(folder_info, columns=['Folder Path', 'Folder Name', 'File Name', 'File Size(KB)'])
        df.to_excel(writer, index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        bold_format = workbook.add_format({'bold': True, 'align': 'center'})
        center_format = workbook.add_format({'align': 'center'})
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, bold_format)
            worksheet.set_column(col_num, col_num, len(str(value)) + 20)
            worksheet.set_row(0, None, bold_format)
        for row_num, row_data in enumerate(df.values):
            for col_num, value in enumerate(row_data):
                worksheet.write(row_num + 1, col_num, value, center_format)

    # 显示生成结果
    catalogue_result_label.config(text="The catalogue list was output to the file of Catalogue List.xlsx")

# 创建人机界面
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Folder_Check(Made by PP2)")
    root.geometry("800x400")

    # 创建输入框
    folder_path_var = tk.StringVar()
    folder_path_entry = tk.Entry(root, textvariable=folder_path_var, width=50)
    folder_path_entry.pack(pady=10)

    # 创建生成目录列表按钮
    catalogue_button = tk.Button(root, text="File Catalogue", command=generate_catalogue)
    catalogue_button.pack(pady=10)

    # 创建进度条和进度标签
    # progressbar_label = tk.Label(root)
    # progressbar_label.pack()
    # progressbar = ttk.Progressbar(root, orient='horizontal', length=300, mode='determinate')
    # progressbar.pack(pady=10)

    # 创建结果标签和作者
    result_label = tk.Label(root)
    result_label.pack()
    author_label = tk.Label(root)
    author_label.pack()
    catalogue_result_label = tk.Label(root)
    catalogue_result_label.pack()

    # 显示人机界面
    root.mainloop()
