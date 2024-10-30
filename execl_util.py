import os
import shutil

from openpyxl import load_workbook, Workbook


def split_excel(file_path, output_dir, chunk_size):
    # 读取 Excel 文件
    wb = load_workbook(file_path)
    ws = wb.active

    # 获取行数和列数
    total_rows = ws.max_row
    total_cols = ws.max_column

    # 创建表头
    headers = [ws.cell(1, col).value for col in range(1, total_cols + 1)]

    # 如果输出目录存在，先清空目录
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    # 重新创建输出目录
    os.makedirs(output_dir)

    # 按 chunk_size 拆分数据
    for i in range(0, total_rows - 1, chunk_size):
        new_wb = Workbook()
        new_ws = new_wb.active

        # 写入表头
        new_ws.append(headers)

        # 写入每块数据
        for row in range(i + 2, min(i + chunk_size + 2, total_rows + 1)):
            new_row = [ws.cell(row, col).value for col in range(1, total_cols + 1)]
            new_ws.append(new_row)

        # 保存文件到指定目录
        output_file = os.path.join(output_dir, f'output_part_{i // chunk_size + 1}.xlsx')
        new_wb.save(output_file)
        print(f'Saved: {output_file}')


def main():
    # 输入文件路径和拆分行数
    file_path = 'douyin_26.xlsx'  # 替换为你的Excel文件路径
    output_dir = 'execl_directory'  # 替换为你希望保存文件的输出目录
    chunk_size = 199  # 每个新文件的行数，可以根据需求更改

    # 调用拆分函数
    split_excel(file_path, output_dir, chunk_size)


if __name__ == '__main__':
    main()
