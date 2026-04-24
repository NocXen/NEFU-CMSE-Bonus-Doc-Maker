import os
import pandas as pd
import re
from datetime import datetime
import logging

"""
通用Excel重复数据检查插件
功能：检查指定列中去掉空格和换行后的重复内容
"""

# --------------------------------------------------
# 配置更改区
# --------------------------------------------------

# 要检查的列名（可以多个）
COLUMNS_TO_CHECK = ['班级', '姓名', '事由', '性质']

# 输入文件目录（包含要检查的xlsx文件）
INPUT_DIRECTORY = './Input_输入'

# 输出日志目录和文件名
LOG_FILE_PATH = './Checker.log'

# 是否包含子目录
INCLUDE_SUBDIRS = False

# 文件编码（默认为utf-8）
FILE_ENCODING = 'utf-8'

# --------------------------------------------------
# 函数实现区
# --------------------------------------------------

def setup_logging():
    """设置日志配置"""
    # 创建日志目录
    log_dir = os.path.dirname(LOG_FILE_PATH)
    if log_dir and not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    # 配置日志
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(LOG_FILE_PATH, encoding=FILE_ENCODING),
            logging.StreamHandler()  # 同时输出到控制台
        ]
    )

def clean_text(text):
    """清理文本：去掉空格、换行等空白字符"""
    if pd.isna(text) or text is None:
        return ""
    
    # 转换为字符串并清理空白字符
    text_str = str(text).strip()
    # 去掉所有空白字符（空格、换行、制表符等）
    cleaned = re.sub(r'\s+', '', text_str)
    return cleaned

def find_excel_files(directory, include_subdirs=True):
    """查找目录中的所有Excel文件"""
    excel_files = []
    
    if include_subdirs:
        for root, dirs, files in os.walk(directory):
            for file in files:
                if file.lower().endswith('.xlsx'):
                    excel_files.append(os.path.join(root, file))
    else:
        for file in os.listdir(directory):
            if file.lower().endswith('.xlsx') and os.path.isfile(os.path.join(directory, file)):
                excel_files.append(os.path.join(directory, file))
    
    return excel_files

def check_duplicates_in_file(file_path, columns):
    """检查单个文件中的重复数据"""
    duplicates_found = []
    
    try:
        # 读取Excel文件的所有sheet
        excel_file = pd.ExcelFile(file_path)
        
        for sheet_name in excel_file.sheet_names:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                
                # 检查所需的列是否存在
                missing_columns = [col for col in columns if col not in df.columns]
                if missing_columns:
                    logging.warning(f"文件 {file_path} 的工作表 '{sheet_name}' 中缺少列: {missing_columns}")
                    continue
                
                # 清理数据
                cleaned_data = []
                for _, row in df.iterrows():
                    cleaned_row = {col: clean_text(row[col]) for col in columns}
                    cleaned_row['_original_index'] = _  # 保存原始行索引
                    cleaned_data.append(cleaned_row)
                
                # 查找重复项
                seen = {}
                for i, row_data in enumerate(cleaned_data):
                    # 创建行的唯一标识（基于指定列的组合）
                    row_key = tuple(row_data[col] for col in columns)
                    
                    # 跳过空行
                    if all(value == "" for value in row_key):
                        continue
                    
                    if row_key in seen:
                        # 找到重复
                        original_row_index = seen[row_key]['original_index'] + 2  # Excel行号（从1开始，加上标题行）
                        current_row_index = i + 2
                        
                        duplicate_info = {
                            'file_name': os.path.basename(file_path),
                            'file_path': file_path,
                            'sheet_name': sheet_name,
                            'columns': columns.copy(),
                            'duplicate_value': {col: row_data[col] for col in columns},
                            'original_row': original_row_index,
                            'duplicate_row': current_row_index,
                            'all_duplicate_rows': seen[row_key]['rows'] + [current_row_index]
                        }
                        
                        # 更新已记录的行号列表
                        seen[row_key]['rows'].append(current_row_index)
                        duplicates_found.append(duplicate_info)
                        
                    else:
                        # 第一次见到这个值
                        seen[row_key] = {
                            'original_index': i,
                            'rows': [i + 2]  # Excel行号
                        }
                
            except Exception as e:
                logging.error(f"处理文件 {file_path} 的工作表 '{sheet_name}' 时出错: {str(e)}")
                continue
                
    except Exception as e:
        logging.error(f"读取文件 {file_path} 时出错: {str(e)}")
    
    return duplicates_found

def main():
    """主函数"""
    print("开始检查重复数据...")
    print(f"检查列: {COLUMNS_TO_CHECK}")
    print(f"输入目录: {INPUT_DIRECTORY}")
    print(f"日志文件: {LOG_FILE_PATH}")
    print("-" * 40)
    
    # 设置日志
    setup_logging()
    
    # 记录开始时间
    start_time = datetime.now()
    logging.info(f"开始重复数据检查 - {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    logging.info(f"检查列: {COLUMNS_TO_CHECK}")
    logging.info(f"输入目录: {INPUT_DIRECTORY}")
    
    # 检查输入目录是否存在
    if not os.path.exists(INPUT_DIRECTORY):
        logging.error(f"输入目录不存在: {INPUT_DIRECTORY}")
        return
    
    # 查找Excel文件
    excel_files = find_excel_files(INPUT_DIRECTORY, INCLUDE_SUBDIRS)
    
    if not excel_files:
        logging.warning("在指定目录中未找到任何xlsx文件")
        return
    
    logging.info(f"找到 {len(excel_files)} 个Excel文件")
    
    total_duplicates = 0
    
    # 检查每个文件
    for file_path in excel_files:
        logging.info(f"正在检查文件: {file_path}")
        duplicates = check_duplicates_in_file(file_path, COLUMNS_TO_CHECK)
        
        if duplicates:
            total_duplicates += len(duplicates)
            for dup in duplicates:
                # 打印到控制台
                print(f"\n发现重复数据:")
                print(f"  文件: {dup['file_name']}")
                print(f"  工作表: {dup['sheet_name']}")
                print(f"  列: {dup['columns']}")
                print(f"  重复值: {dup['duplicate_value']}")
                print(f"  重复行号: {dup['all_duplicate_rows']}")
                print("-" * 30)
                
                # 记录到日志文件
                logging.info(f"发现重复数据 - 文件: {dup['file_name']}, "
                           f"工作表: {dup['sheet_name']}, "
                           f"列: {dup['columns']}, "
                           f"重复值: {dup['duplicate_value']}, "
                           f"重复行号: {dup['all_duplicate_rows']}")
    
    # 记录结束时间和统计信息
    end_time = datetime.now()
    duration = end_time - start_time
    
    summary_msg = f"检查完成。共检查 {len(excel_files)} 个文件，发现 {total_duplicates} 处重复数据。耗时: {duration}"
    logging.info(summary_msg)
    print(f"\n{summary_msg}")

if __name__ == "__main__":
    main()
    input("\n按回车键退出...\n")