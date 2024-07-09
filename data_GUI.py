import re
import pandas as pd
import tkinter as tk
import numpy as np
import tkinter.font as tkFont
from tkinter import messagebox, scrolledtext
import shutil
import os
import glob
from datetime import datetime

data_path = "data_set.xlsx"

def backup_excel_file(original_file, max_backups=10):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = original_file.replace('.xlsx', f'_backup_{timestamp}.xlsx')
    shutil.copyfile(original_file, backup_file)

    # 获取所有备份文件并按修改时间排序
    backup_files = sorted(glob.glob(original_file.replace('.xlsx', '_backup_*.xlsx')), key=os.path.getmtime)

    # 保留最新的max_backups个备份，删除其余的
    for old_backup in backup_files[:-max_backups]:
        os.remove(old_backup)

# 合金自动识别和处理
def parse_alloy_composition(alloy_str):
    # 去除空格并初始化字典
    alloy_str = alloy_str.replace(" ", "")
    element_counts = {}

    # 正则表达式分离括号内外的元素
    pattern = r'\((?P<inner>[A-Za-z\d\.]+)\)(?P<inner_multiplier>\d*)|(?P<element>[A-Z][a-z]*)(?P<count>\d*\.?\d*)'
    matches = re.findall(pattern, alloy_str)

    # 括号内外的总量计算
    inner_multiplier = 1  # 默认为1，处理括号外无数字的情况
    outer_total = 0
    inner_composition = ""

    for inner, inner_mult, element, count in matches:
        if inner:
            inner_composition = inner
            inner_multiplier = float(inner_mult) if inner_mult else 1  # 处理括号外无数字的情况
        else:
            count = float(count) if count else 1
            outer_total += count
            element_counts[element] = count

    # 计算括号内元素的比例
    inner_matches = re.findall(r'([A-Z][a-z]*)(\d*\.?\d*)', inner_composition)
    inner_total = sum([float(c) if c else 1 for _, c in inner_matches])

    for e, c in inner_matches:
        c = float(c) if c else 1
        element_proportion = c / inner_total
        element_counts[e] = element_counts.get(e, 0) + element_proportion * inner_multiplier

    # 调整所有元素的比例
    total_amount = sum(element_counts.values())
    for element in element_counts:
        element_counts[element] = round((element_counts[element] / total_amount) * 100, 4)

    return element_counts

# 性能输入与检查
def process_property_input(property_value):
    if property_value.strip() == "":
        return None
    try:
        return float(property_value)
    except ValueError:
        return "错误：请输入有效的数字。"

# 相成分输入
def process_phase_properties(phase_selections):
    phases = ["FCC", "BCC", "L12", "B2", "L21", "Laves", "σ", "Im"]
    phase_properties = {}

    # 检查是否有任何"是"的选择
    any_yes = any(choice == '是' for choice in phase_selections.values())

    for phase in phases:
        if any_yes:
            # 对于新合金，如果有"是"的选择，则设置为1或0
            phase_properties[phase] = 1 if phase_selections.get(phase) == '是' else 0
        else:
            # 如果没有"是"的选择，则保留为空
            phase_properties[phase] = None

    return phase_properties


# 输入备注
def process_custom_properties(eutectic_input, doi_input):
    return {
        "共晶": eutectic_input if eutectic_input.strip() != "" else None,
        "Doi": doi_input if doi_input.strip() != "" else None
    }

# 检查是否存在相同成分的合金
def check_if_alloy_exists(df, alloy_composition_parsed):

    alloy_exists = df[df[list(alloy_composition_parsed.keys())].eq(pd.Series(alloy_composition_parsed)).all(axis=1)]
    return not alloy_exists.empty, alloy_exists

# 读取函数
def read_excel(file_path):
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        print(f"读取文件时发生错误: {e}")
        return None

# 定义数据条数标签
def update_data_count_label(df, label):
    if df is not None:
        count = len(df)
        label.config(text=f"数据条数: {count}")
    else:
        label.config(text="数据条数: 0")


# 写入函数
def write_excel(df, file_path):
    # 先进行备份
    backup_excel_file(file_path)

    # 然后写入数据
    try:
        df.to_excel(file_path, index=False)
        print("文件已成功保存")
    except Exception as e:
        print(f"写入文件时发生错误: {e}")

    # 更新数据条数显示
    update_data_count_label(df, data_count_label)

# 用于打开一个新窗口让用户输入性能信息
def open_new_window_for_input(alloy_data, alloy_composition_parsed, result_area, alloy_entry, new_entry=True):
    input_window = tk.Toplevel()
    input_window.title("输入性能信息")

    # 定义字体
    chinese_font = tkFont.Font(family="楷体", size=14, weight="bold")
    english_font = tkFont.Font(family="Times New Roman", size=14, weight="bold")

    # 为性能指标创建输入框
    properties = ["T_YS", "T_US", "T_Elo", "C_YS", "C_Elo", "Hardness", "Ecorr", "Epit"]
    entries = {}
    for idx, prop in enumerate(properties):
        tk.Label(input_window, text=prop + ":", font=chinese_font).grid(row=idx, column=0, sticky='w', padx=10, pady=5)
        entry = tk.Entry(input_window, font=english_font)
        entry.grid(row=idx, column=1, padx=10, pady=5)
        entries[prop] = entry

    # 添加FCC, BCC等属性的按钮
    phase_vars = {}
    phases = ["FCC", "BCC", "L12", "B2", "L21", "Laves", "σ", "Im"]
    for idx, phase in enumerate(phases, start=len(properties)):
        tk.Label(input_window, text=phase + ":", font=chinese_font).grid(row=idx, column=0, sticky='w', padx=10, pady=5)
        var = tk.StringVar(value='未选择' if new_entry else None)
        yes_button = tk.Radiobutton(input_window, text="是", variable=var, value='是', font=chinese_font)
        no_button = tk.Radiobutton(input_window, text="否", variable=var, value='否', font=chinese_font)
        none_button = tk.Radiobutton(input_window, text="未选择", variable=var, value='未选择', font=chinese_font)
        yes_button.grid(row=idx, column=1, padx=10, pady=5)
        no_button.grid(row=idx, column=2, padx=10, pady=5)
        none_button.grid(row=idx, column=3, padx=10, pady=5)
        phase_vars[phase] = var

    # 添加共晶和Doi的输入框
    eutectic_entry = tk.Entry(input_window, font=english_font)
    doi_entry = tk.Entry(input_window, font=english_font)
    eutectic_entry.grid(row=len(phases) + len(properties), column=1, padx=10, pady=5)
    doi_entry.grid(row=len(phases) + len(properties) + 1, column=1, padx=10, pady=5)
    tk.Label(input_window, text="共晶:", font=chinese_font).grid(row=len(phases) + len(properties), column=0, sticky='w', padx=10, pady=5)
    tk.Label(input_window, text="Doi:", font=chinese_font).grid(row=len(phases) + len(properties) + 1, column=0, sticky='w', padx=10, pady=5)

    # 提交按钮
    submit_button = tk.Button(input_window, text="提交", font=chinese_font, command=lambda: submit_performance_data(
        entries, alloy_data, alloy_composition_parsed, result_area, input_window, alloy_entry, phase_vars,
        eutectic_entry, doi_entry))
    submit_button.grid(row=len(phases) + len(properties) + 2, column=1, padx=10, pady=5)


# 更新信息窗口
def save_updated_alloy_data(entries, phase_vars, eutectic_entry, doi_entry, existing_alloys, result_area, update_window, alloy_entry):
    global df
    alloy_index = existing_alloys.index[0]  # 假设只处理第一个匹配的合金

    # 提取并更新性能数据
    for prop, entry in entries.items():
        new_value = process_property_input(entry.get())
        if new_value != existing_alloys.at[alloy_index, prop]:
            df.loc[alloy_index, prop] = new_value

    # 更新相位属性
    for phase, var in phase_vars.items():
        df.loc[alloy_index, phase] = 1 if var.get() == '1' else 0

    # 更新共晶和Doi
    df.loc[alloy_index, "共晶"] = eutectic_entry.get() if eutectic_entry.get() else df.loc[alloy_index, "共晶"]
    df.loc[alloy_index, "Doi"] = doi_entry.get() if doi_entry.get() else df.loc[alloy_index, "Doi"]

    # 保存更新到Excel
    write_excel(df, data_path)

    # 重新加载数据集
    df = read_excel(data_path)

    # 更新主窗口结果区域
    updated_info = df.loc[alloy_index].to_string()
    result_area.delete('1.0', tk.END)
    result_area.insert(tk.INSERT, "更新后的合金数据:\n" + updated_info)

    # 清空合金成分输入框
    alloy_entry.delete(0, tk.END)

    # 关闭更新窗口
    update_window.destroy()


# 重复值处理窗口
def open_window_for_existing_alloy(alloy_data, existing_alloys, result_area, alloy_entry):
    update_window = tk.Toplevel()
    update_window.title("更新合金信息")

    # 显示现有合金信息
    existing_info = existing_alloys.iloc[0]  # 假设只处理一个重复的合金
    row_idx = 0
    # 定义字体
    chinese_font = tkFont.Font(family="楷体", size=14, weight="bold")
    english_font = tkFont.Font(family="Times New Roman", size=14, weight="bold")

    # 创建性能指标输入框并初始化为现有值
    entries = {}
    for prop in ["T_YS", "T_US", "T_Elo", "C_YS", "C_Elo", "Hardness", "Ecorr", "Epit"]:
        tk.Label(update_window, text=prop + ":", font=chinese_font).grid(row=row_idx, column=0, sticky='w', padx=10,
                                                                         pady=5)
        entry = tk.Entry(update_window, font=english_font)
        entry.insert(0, existing_info[prop])  # 初始化为现有值
        entry.grid(row=row_idx, column=1)
        entries[prop] = entry
        row_idx += 1

    # 创建相位属性的单选按钮并初始化为现有值
    phase_vars = {}
    for phase in ["FCC", "BCC", "L12", "B2", "L21", "Laves", "σ", "Im"]:
        tk.Label(update_window, text=phase + ":", font=chinese_font).grid(row=row_idx, column=0, sticky='w', padx=10,
                                                                          pady=5)
        # 检查NaN值并相应地处理
        if np.isnan(existing_info[phase]):
            phase_value = '未选择'
        else:
            phase_value = str(int(existing_info[phase]))

        var = tk.StringVar(value=phase_value)
        yes_button = tk.Radiobutton(update_window, text="是", variable=var, value='1', font=chinese_font)
        no_button = tk.Radiobutton(update_window, text="否", variable=var, value='0', font=chinese_font)

        yes_button.grid(row=row_idx, column=1)
        no_button.grid(row=row_idx, column=2)
        phase_vars[phase] = var
        row_idx += 1

    # 创建共晶和Doi的输入框并初始化为现有值
    eutectic_entry = tk.Entry(update_window, font=english_font)
    doi_entry = tk.Entry(update_window, font=english_font)
    eutectic_entry.insert(0, existing_info["共晶"])
    doi_entry.insert(0, existing_info["Doi"])
    eutectic_entry.grid(row=row_idx, column=1, padx=10, pady=5)
    doi_entry.grid(row=row_idx + 1, column=1, padx=10, pady=5)
    tk.Label(update_window, text="共晶:", font=chinese_font).grid(row=row_idx, column=0, sticky='w', padx=10, pady=5)
    tk.Label(update_window, text="Doi:", font=chinese_font).grid(row=row_idx + 1, column=0, sticky='w', padx=10, pady=5)

    # 保存按钮
    save_button = tk.Button(update_window, text="保存",
                            command=lambda: save_updated_alloy_data(
                                entries, phase_vars, eutectic_entry, doi_entry, existing_alloys, result_area,
                                update_window, alloy_entry))
    save_button.grid(row=row_idx + 2, column=1)


def submit_performance_data(entries, alloy_data, alloy_composition_parsed, result_area, input_window, alloy_entry, phase_vars, eutectic_entry, doi_entry):
    global df
    performance_data = {}
    valid_input = True

    for prop, entry in entries.items():
        property_value = process_property_input(entry.get())
        if isinstance(property_value, str):  # 检查是否返回错误消息
            messagebox.showerror("输入错误", f"{prop}: {property_value}")
            valid_input = False
            break
        performance_data[prop] = property_value

    if valid_input:

        # 确保所有未提及的元素设置为0
        all_elements = set(
            ['Fe', 'Ni', 'Cr', 'Al', 'Ti', 'Mo', 'Nb', 'V', 'Co', 'Mn', 'Cu', 'Zr', 'Hf', 'Ta', 'Si', 'W'])
        for element in all_elements - set(alloy_composition_parsed.keys()):
            alloy_composition_parsed[element] = 0

        # 处理相位属性输入
        phase_selections = {phase: var.get() for phase, var in phase_vars.items()}
        phase_properties = process_phase_properties(phase_selections)

        # 处理共晶和Doi输入
        custom_properties = process_custom_properties(eutectic_entry.get(), doi_entry.get())

        # 整合所有数据
        alloy_data_row = {"Alloy": alloy_data}
        alloy_data_row.update(alloy_composition_parsed)
        alloy_data_row.update(performance_data)
        alloy_data_row.update(phase_properties)
        alloy_data_row.update(custom_properties)

        # 将数据写入Excel表格
        df = read_excel(data_path)  # 读取现有数据
        new_df = pd.DataFrame([alloy_data_row])  # 创建新数据行
        updated_df = pd.concat([df, new_df], ignore_index=True)
        write_excel(updated_df, data_path)  # 写入更新后的数据

        # 重新加载数据集
        df = read_excel(data_path)

        # 清空合金成分输入框
        alloy_entry.delete(0, tk.END)

        # 关闭输入窗口
        input_window.destroy()

        # 更新主窗口结果区域
        result_area.delete('1.0', tk.END)
        result_str = "\n".join(f"{k}: {v}" for k, v in alloy_data_row.items())
        result_area.insert(tk.INSERT, "保存的合金数据:\n" + result_str)

        # 关闭输入窗口
        input_window.destroy()

def main():
    global data_count_label  # 声明为全局变量

    # 创建主窗口
    root = tk.Tk()
    root.title("合金成分和性能输入")
    root.geometry("800x850")

    # 定义字体
    chinese_font = tkFont.Font(family="楷体", size=14, weight="bold")
    english_font = tkFont.Font(family="Times New Roman", size=14, weight="bold")

    # 合金成分输入标签和输入框
    label_alloy = tk.Label(root, text="合金成分:", font=chinese_font)
    label_alloy.grid(row=0, column=0, padx=10, pady=10)
    alloy_entry = tk.Entry(root, font=english_font, width=50)
    alloy_entry.grid(row=0, column=1, padx=10, pady=10)

    # 结果显示区域
    result_area = scrolledtext.ScrolledText(root, height=30, width=70, font=english_font)
    result_area.grid(row=2, column=0, columnspan=2, padx=10, pady=10)

    # 提交按钮
    submit_button = tk.Button(root, text="确定", font=chinese_font,
                              command=lambda: submit_alloy_data(alloy_entry.get(), result_area, alloy_entry))
    submit_button.grid(row=1, column=0, columnspan=2, padx=10, pady=10)

    # 数据条数标签
    data_count_label = tk.Label(root, text="数据条数: 0", font=chinese_font)
    data_count_label.grid(row=3, column=0, columnspan=2, sticky='ew', padx=10, pady=5)

    # 加载数据集并更新标签
    global df
    df = read_excel(data_path)
    update_data_count_label(df, data_count_label)

    root.mainloop()

def dataframe_to_formatted_string(df):
    formatted_str = ""
    for index, row in df.iterrows():
        row_str = f"行 {index}:\n"
        for col in df.columns:
            row_str += f"    {col}: {row[col]}\n"
        formatted_str += row_str + "\n"
    return formatted_str

def submit_alloy_data(alloy_data, result_area, alloy_entry):
    alloy_composition_parsed = parse_alloy_composition(alloy_data)
    exists, existing_alloys = check_if_alloy_exists(df, alloy_composition_parsed)

    if exists:
        existing_info = dataframe_to_formatted_string(existing_alloys)
        result_area.delete('1.0', tk.END)
        result_area.insert(tk.INSERT, f"重复合金信息:\n{existing_info}")
        # 打开更新窗口
        open_window_for_existing_alloy(alloy_data, existing_alloys, result_area, alloy_entry)
    else:
        open_new_window_for_input(alloy_data, alloy_composition_parsed, result_area, alloy_entry)
if __name__ == "__main__":
    main()
