import os
from datetime import datetime
import pandas as pd
from tkinter import messagebox
import tkinter as tk

def perform_merge(folder_path):
    # 获取当前运行脚本的目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    # 向上两级目录找到项目根目录（假设 main.py 在项目根目录下）
    project_root = os.path.dirname(os.path.dirname(current_dir))
    project_root = os.path.join(project_root, 'File_selector_v2.0')
    current_time = datetime.now().strftime('%Y%m%d%H%M%S')
    output_folder = f"{current_time}OutputFolder"
    output_folder_path = os.path.join(project_root, output_folder)
    os.makedirs(output_folder_path, exist_ok=True)

    all_data = []
    file_count = 0
    for root, dirs, files in os.walk(folder_path):
        file_pairs = [files[i:i + 2] for i in range(0, len(files), 2)]
        for pair in file_pairs:
            if len(pair) == 2:
                data_list = []
                for file in pair:
                    if file.endswith('.txt'):
                        file_path = os.path.join(root, file)
                        with open(file_path, 'r', encoding='utf-8') as txt_file:
                            lines = txt_file.readlines()
                            found_keyword = False
                            for index, line in enumerate(lines):
                                if line.startswith("mm\tkN\tsec"):
                                    found_keyword = True
                                    continue
                                if found_keyword:
                                    data = line.split('\t')[:2]
                                    if index >= 1 and index % 200 == 0:
                                        data_list.append(['', ''])
                                    # 过滤掉空字符串
                                    if all(item.strip() for item in data):
                                        data_list.append(data)

                # 对这对文件的数据进行处理
                processed_data = []
                for row in data_list:
                    if len(row) == 2:
                        try:
                            first_value = float(row[0]) / 73 if row[0].strip() else 0
                            second_value = (float(row[1]) / 70.7 / 70.7 * 1000) if row[1].strip() else 0
                            processed_data.append([first_value, second_value])
                        except ValueError:
                            processed_data.append([0, 0])

                df = pd.DataFrame(processed_data, columns=[f"Column1_{file_count}", f"Column2_{file_count}"])
                all_data.append(df)
                file_count += 1

    combined_df = pd.concat(all_data, axis=1)

    new_file_name = os.path.join(output_folder_path, f"{current_time}_combined.xlsx")
    combined_df.to_excel(new_file_name, index=False)

    # 弹出成功消息框并显示文件保存位置
    message = f"文件合并已完成！\n文件保存位置：{new_file_name}"
    tk.messagebox.showinfo("合并成功", message)