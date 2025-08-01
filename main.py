# 文件名：process_excel.py
from pathlib import Path

import pandas as pd

print("-------------------------------")
print("注意事项：")
print("1. 无论是什么东西，命名不建议带空格，否则可能出bug,建议用空格连接")
print("2. 建议把这个脚本和你要处理的文件放在同个目录下执行，这样你只输入文件名字就行。windows的路径是很难搞对的……")
print("3. 最终解释权归本观测员所有")
print("-------------------------------")
print(f"\n\n\n")
# file_name = "FILTER.xlsx"
# workspace_sheet = "OD"
# condition_sheet = "CL"
file_name = input("请输入文件名称（如不在同一路径下，请输入全路径）：")

if file_name == "":
    exit("输入的文件名不能为空")

if not file_name.endswith(".xlsx"):
    exit("仅支持.xlsx格式的文件")

if not Path(file_name).exists():
    exit(f"文件{file_name}不存在")

workspace_sheet = input("请输入需要筛选的sheet名称：")
condition_sheet = input("请输入条件sheet名称：")

header_value = input("请输入要匹配的表头(第一行的值)，比如A1那一格：")

# 读取 Excel 文件
try:
    ws = pd.read_excel(file_name, sheet_name=workspace_sheet)
except Exception as e:
    exit(f"读取工作表{workspace_sheet}出错：{e}")

try:
    cs = pd.read_excel(file_name, sheet_name=condition_sheet)
except Exception as e:
    exit(f"读取工作表{condition_sheet}出错：{e}")

# -----------------------------
# 遍历读取条件
# 山哥要的多条件
# -----------------------------
all_cells = [cs.columns.tolist()]  # 表头作为第一行
all_cells.extend(cs.values.tolist())  # 所有数据行添加到列表
conditions: list[str] = []
for row_idx, row in enumerate(all_cells):
    for col_idx, value in enumerate(row):
        if value != "" and value is not None:
            conditions.append(str(value))
        # print(f"第{row_idx+1}行，第{col_idx+1}列：{value}")

if len(conditions) == 0:
    exit(f"没有条件喔")

# 过滤掉所有全为空值的列（包括 Unnamed 列）
ws = ws.dropna(axis=1, how='all')


# 精确匹配A列中等于输入值的所有行
# 使用str.fullmatch确保完全匹配，避免部分包含的情况
matches = pd.DataFrame()
for item in conditions:
    print(f"-----------------------{item}")
    single_match = ws[ws[header_value].str.fullmatch(item, na=False)]

    # 将当前匹配结果合并到总结果中
    matches = pd.concat([matches, single_match], ignore_index=True)

# 输出结果
print(f"\n{header_value}列中值为'{conditions}'的所有行：")
if not matches.empty:
    print(matches)
else:
    print("没有找到匹配的记录")

# 可选：将结果保存到新sheet
sheet_save = input("\n将结果保存到新表(sheet)？(y/n)：")
if sheet_save.lower() == 'y':
    filtered_sheet_name = input("\n请输入新的表名：")
    with pd.ExcelWriter(file_name, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        matches.to_excel(writer, sheet_name=filtered_sheet_name, index=False)
        print(f"结果已保存到{filtered_sheet_name}")

# 可选：将结果保存到新文件
file_save = input("\n将结果保存到新的Excel文件？(y/n)：")
if file_save.lower() == 'y':
    matches.to_excel(f"{header_value}列匹配.xlsx", index=False)
    print(f"结果已保存为：{header_value}列匹配.xlsx")
