import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment
#
# 设置误差范围
tolerance = 0.0030
normal_df = pd.read_excel('data/normal.xlsx')
cnacer_df = pd.read_excel('data/cnacer.xlsx')

# 提取第一列 m/z 数据
normal_mz = normal_df.iloc[:, 0]
cnacer_mz = cnacer_df.iloc[:, 0]

# 取第一行作为列名
normal_cols = normal_df.columns.tolist()
cnacer_cols = cnacer_df.columns.tolist()

# 初始化结果列表
matched_rows = []
remaining_normal = []
remaining_cnacer = []

# 找到匹配项
for i, mz_normal in normal_mz.items():
    found_match = False
    for j, mz_cnacer in cnacer_mz.items():
        if abs(mz_normal - mz_cnacer) <= tolerance:
            # 将 normal 和 cnacer 的数据合并在一起
            combined_row = pd.concat([normal_df.iloc[i], cnacer_df.iloc[j]], ignore_index=True)
            combined_row = combined_row.rename(lambda x: f"C{x}")  # 暂时使用 C0、C1 等命名列
            combined_row['m/z'] = mz_normal  # 添加 m/z 列方便分组
            matched_rows.append(combined_row)
            found_match = True
            break
    if not found_match:
        # 如果没有匹配项，加入到 remaining_normal 中
        remaining_normal.append(normal_df.iloc[i])

# 查找剩余的 cnacer 项
for j, mz_cnacer in cnacer_mz.items():
    if not any(abs(mz_cnacer - mz_normal) <= tolerance for mz_normal in normal_mz):
        remaining_cnacer.append(cnacer_df.iloc[j])

# 创建 DataFrame 并设置合适的列名
if matched_rows:
    matched_df = pd.DataFrame(matched_rows)
    matched_df.sort_values(by='m/z', inplace=True)  # 按 m/z 排序

    # 使用 normal 和 cnacer 的列名
    column_names = normal_cols + cnacer_cols
    matched_df.columns = column_names + ["m/z"]  # 添加 "m/z" 列用于后续操作

    # 删除最后的 "m/z" 列
    matched_df.drop(columns=["m/z"], inplace=True)

# 将结果写入 Excel 文件
with pd.ExcelWriter('data/matched.xlsx', engine='openpyxl') as writer:
    if not matched_df.empty:
        matched_df.to_excel(writer, index=False, sheet_name='Matched')
    else:
        pd.DataFrame().to_excel(writer, index=False, sheet_name='Matched')  # 空文件

# 加载已保存的文件以进行单元格合并
wb = load_workbook('data/matched.xlsx')
ws = wb['Matched']

# 合并相同 m/z 值的单元格
for row in range(2, ws.max_row + 1):
    if ws[f'A{row}'].value == ws[f'A{row - 1}'].value:
        ws.merge_cells(f'A{row - 1}:A{row}')
        ws[f'A{row - 1}'].alignment = Alignment(horizontal="center", vertical="center")

wb.save('data/matched.xlsx')

# 生成 remaining 文件
with pd.ExcelWriter('data/remaining_normal.xlsx') as writer:
    if remaining_normal:
        pd.DataFrame(remaining_normal).to_excel(writer, index=False, sheet_name='Remaining_Normal')
    else:
        pd.DataFrame().to_excel(writer, index=False, sheet_name='Remaining_Normal')  # 空文件

with pd.ExcelWriter('data/remaining_cnacer.xlsx') as writer:
    if remaining_cnacer:
        pd.DataFrame(remaining_cnacer).to_excel(writer, index=False, sheet_name='Remaining_Cnacer')
    else:
        pd.DataFrame().to_excel(writer, index=False, sheet_name='Remaining_Cnacer')  # 空文件

print("文件对比完成。生成的文件分别为 matched.xlsx、remaining_normal.xlsx、remaining_cnacer.xlsx")