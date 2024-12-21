import pandas as pd
import os

# 基础目录
base_dir = 'D:/Tile/Work Directory/To-do/BVF'
csv_dir = os.path.join(base_dir, 'csv')

# 创建输出文件夹
output_dir = os.path.join(base_dir, 'output')
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# 文件路径
file_path = os.path.join(csv_dir, '8.29.csv')
output_file = os.path.join(output_dir, 'output.txt')
in_file = os.path.join(csv_dir, 'in.csv')
out_file = os.path.join(csv_dir, 'out.csv')
result_txt = os.path.join(output_dir, 'result.txt')
find_txt = os.path.join(output_dir, 'find.txt')
result_excel = os.path.join(output_dir, 'result.xlsx')

try:
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"文件 {file_path} 不存在，请检查路径。")
    
    df = pd.read_csv(file_path, encoding='utf-8', sep=',', header=None)
    df = df.drop(index=0).reset_index(drop=True)
    df.to_csv(output_file, sep='|', index=False, header=False, encoding='utf-8')

    in_names = set(pd.read_csv(in_file, encoding='utf-8', sep=',').iloc[:, 0].astype(str))
    out_names = set(pd.read_csv(out_file, encoding='utf-8', sep=',').iloc[:, 0].astype(str))

    with open(output_file, 'r', encoding='utf-8') as f:
        result_txt_lines = []
        find_txt_lines = []
        for line in f:
            name = line.split('|')[0]
            if name in in_names or name in out_names:
                find_txt_lines.append(line)
            else:
                result_txt_lines.append(line)

    with open(result_txt, 'w', encoding='utf-8') as txt_file:
        txt_file.writelines(result_txt_lines)

    with open(find_txt, 'w', encoding='utf-8') as txt_file:
        txt_file.writelines(find_txt_lines)

    os.remove(output_file)

    result_df = pd.read_csv(result_txt, sep='|', header=None, encoding='utf-8')
    result_df.to_excel(result_excel, index=False, header=False)
    print(f"数据已成功写入")

except FileNotFoundError as e:
    print(e)
except Exception as e:
    print(f"读取文件时出错: {e}")