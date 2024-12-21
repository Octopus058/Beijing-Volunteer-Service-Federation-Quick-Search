import os
import pandas as pd

# 确保文件路径正确
input_dir = os.path.join(os.getcwd(), 'csv', 'raw')
output_dir = os.path.join(os.getcwd(), 'csv', 'name list')

# 确保输出目录存在
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# 需要处理的文件列表
files_to_process = ['in.html', 'out.html']

# 遍历需要处理的文件
for file_name in files_to_process:
    file_path = os.path.join(input_dir, file_name)
    
    try:
        # 加载 HTML 文件
        df = pd.read_html(file_path)[0]
        
        # 只保留 D 列
        df = df.iloc[:, [3]]
        
        # 保存为 CSV 文件
        csv_file_name = f"{os.path.splitext(file_name)[0]}.csv"
        csv_file_path = os.path.join(output_dir, csv_file_name)
        df.to_csv(csv_file_path, index=False, encoding='utf-8-sig')
    
    except Exception as e:
        print(f"处理文件 {file_name} 时出错: {e}")

print("所有文件已成功转换为 CSV 文件。")