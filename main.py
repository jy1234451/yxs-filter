# 文件名：process_excel.py
import pandas as pd

# 读取 Excel 文件
df = pd.read_excel("FILTER.xlsx")


# 过滤掉所有全为空值的列（包括 Unnamed 列）
df = df.dropna(axis=1, how='all')

# 获取用户输入的匹配值
header_value = input("请输入要匹配的表头(第一行的值)，比如A1那一格的值：")
search_value = input(f"请输入要匹配{header_value}列的值：")

# 精确匹配A列中等于输入值的所有行
# 使用str.fullmatch确保完全匹配，避免部分包含的情况
matches = df[df[header_value].str.fullmatch(search_value, na=False)]

# 输出结果
print(f"\n{header_value}列中值为'{search_value}'的所有行：")
if not matches.empty:
    print(matches)
else:
    print("没有找到匹配的记录")

# 可选：将结果保存到新文件
save = input("\n是否将结果保存到Excel？(y/n)：")
if save.lower() == 'y':
    matches.to_excel(f"{header_value}列匹配_{search_value}.xlsx", index=False)
    print(f"结果已保存为：{header_value}列匹配_{search_value}.xlsx")
