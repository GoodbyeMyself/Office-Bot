import os
import pandas as pd

# 定义目录路径
input_folder = 'input'
origin_folder = 'origin'
output_folder = 'out'

# 确保输出目录存在
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# 获取origin目录中的原始Excel文件
origin_files = [f for f in os.listdir(origin_folder) if f.endswith('.xlsx')]

if not origin_files:
    print("origin 目录中没有找到 Excel 文件")
    exit()

# 获取input目录中的Excel文件
input_files = [f for f in os.listdir(input_folder) if f.endswith('.xlsx')]

if not input_files:
    print("input 目录中没有找到 Excel 文件")
    exit()

# 处理每个origin文件
for origin_file in origin_files:
    print(f"正在处理原始文件: {origin_file}")
    
    # 读取原始文件
    origin_path = os.path.join(origin_folder, origin_file)
    origin_df = pd.read_excel(origin_path)
    
    # 初始化数据帧列表，首先添加原始文件数据
    dataframes = [origin_df]
    
    # 读取input目录中的所有Excel文件并添加到列表中
    for input_file in input_files:
        input_path = os.path.join(input_folder, input_file)
        input_df = pd.read_excel(input_path)
        dataframes.append(input_df)
        print(f"  添加文件: {input_file}")
    
    # 合并所有数据帧
    merged_df = pd.concat(dataframes, ignore_index=True)
    
    # 过滤掉空白行（所有列都为空或NaN的行）
    # 记录过滤前的行数
    rows_before_filter = len(merged_df)
    
    # 删除所有列都为空的行
    merged_df = merged_df.dropna(how='all')
    
    # 删除所有列都是空字符串或只包含空白字符的行
    # 首先将所有NaN值替换为空字符串，然后检查是否为空白
    merged_df_str = merged_df.fillna('').astype(str)
    # 创建一个布尔掩码，标识非空白行
    non_empty_mask = merged_df_str.apply(lambda row: row.str.strip().ne('').any(), axis=1)
    merged_df = merged_df[non_empty_mask]
    
    # 记录过滤后的行数
    rows_after_filter = len(merged_df)
    filtered_rows = rows_before_filter - rows_after_filter
    
    # 输出到out目录，使用原始文件名
    output_path = os.path.join(output_folder, origin_file)
    merged_df.to_excel(output_path, index=False)
    
    print(f"已将 {len(input_files)} 个input文件合并到原始文件 {origin_file}")
    print(f"合并结果已保存到: {output_path}")
    print(f"过滤前行数: {rows_before_filter}")
    print(f"过滤掉空白行: {filtered_rows} 行")
    print(f"最终行数: {rows_after_filter}")
    print("-" * 50)

print("所有文件处理完成！")