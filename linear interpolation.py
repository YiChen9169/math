import pandas as pd
import numpy as np

# 读取Excel文件
excel_file = '白金電阻溫度對照表.xlsx'
df = pd.read_excel(excel_file)

# 指定要查找的列
column_R = df['R']    # 第R列（x值）
column_T = df['T']    # 第T列（y值）
column_R_ = df['R_']  # 第R_列（已知的x值）
column_T_ = df['T_']  # 第T_列（用于存储内插结果的y值）

# 定义一个函数，进行线性内插
def linear_interpolation(x, x1, y1, x2, y2):
    return y1 + ((x - x1) / (x2 - x1)) * (y2 - y1)

# 定义一个函数，找到最接近的两个值
def find_closest_values(target_value, values):
    values = sorted(values)  # 对值进行排序
    closest_values = []

    for i in range(len(values) - 1):
        if values[i] <= target_value <= values[i + 1]:
            closest_values = [values[i], values[i + 1]]
            break

    return closest_values

# 逐一遍历R_列的值，并进行线性内插
for i, value_R_ in enumerate(column_R_):
    # 查找最接近的两个R列的值
    closest_values_R = find_closest_values(value_R_, column_R)
    
    if len(closest_values_R) == 2:
        # 获取最接近的两个R列的值和它们对应的T列的值
        x1, x2 = closest_values_R
        y1 = column_T[df['R'] == x1].values[0]
        y2 = column_T[df['R'] == x2].values[0]
        
        # 进行线性内插
        interpolated_y = linear_interpolation(value_R_, x1, y1, x2, y2)
        
        # 将内插结果存储在T_列中
        column_T_[i] = interpolated_y

# 将'T_'列中的值四舍五入到小数点后第2位
df['T_'] = df['T_'].round(2)

# 保存修改后的DataFrame到Excel文件
df.to_excel('白金電阻溫度對照表__.xlsx', index=False)