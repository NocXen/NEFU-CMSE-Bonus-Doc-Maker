import pandas as pd
import os
import re
import datetime
from pathlib import Path
import shutil

"""
填充器 v1.1，
DeepSeek R1 制作。
NocXen 修正。
翻新：2次。

Combiner v1.1 By DeepSeek R1.
Fixed By NocXen.
Rebuilt 2 times.

用途：从输入路径读取月报表xlsx表格，
      将有效信息读取后转换为审批单/回执单信息，
      根据数据文件模板，写入指定数据文件
      （ 定义："-  @ "开头的行为替换行 ）
      ( DR_data_Template.txt -> DR_data.txt )
      ( DocxReplacer的数据文件 )
      并自动启动接下来的: 
      DocxReplacer 与 Combiner 脚本。
      依赖：pandas
"""

# --------------------------------------------------
# 配置和路径设置
# --------------------------------------------------

INPUT_EXCEL_DIR = Path("./Input_输入")
TEMPLATE_DATA_FILE = Path("./Template_模板/DR_data_Template.txt")
OUTPUT_DATA_FILE = Path("./DR_data.txt")
SCRIPT_FILE_1 = Path("./DocxReplacer.py")
SCRIPT_FILE_2 = Path("./Combiner.py")

# --------------------------------------------------
# 班级名称转换函数
# --------------------------------------------------

def convert_class_name(class_name):
    # 提取年份数字
    year_match = re.search(r'(\d{2})', class_name)
    year = f"20{year_match.group(1)}" if year_match else ""
    
    # 专业名称映射
    major_mapping = {
        "林业工程类": "木工",
        "木材科学与工程": "木工",
        "林产化工": "林化",
        "轻化工程": "轻化",
        "材料化学": "材化",
        "高分子材料与工程": "高分子",
        "材料类": "材料类"
    }
    
    # 特殊班级类型处理
    special_classes = {
        "“成栋班”": "成栋",
        "卓越": "卓越"
    }
    
    # 确定专业名称
    major = ""
    for key in major_mapping:
        if key in class_name:
            major = major_mapping[key]
            break
    
    # 处理特殊班级类型
    class_type = ""
    for key in special_classes:
        if key in class_name:
            class_type = special_classes[key]
            break
    
    # 处理普通班级编号
    if not class_type:
        class_num_match = re.search(r'-(\d+)$', class_name)
        if class_num_match:
            num = int(class_num_match.group(1))
            # 数字转中文（支持1-99）
            chinese_nums = ["", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]
            if num <= 10:
                class_type = chinese_nums[num] + "班"
            elif num <= 99:
                tens = num // 10
                units = num % 10
                class_type = (chinese_nums[tens] + "十" if tens > 1 else "十") + chinese_nums[units] + "班"
    
    return f"{year}{major}{class_type}"

# --------------------------------------------------
# 日期处理函数
# --------------------------------------------------

def process_date(date_str):
    """处理日期字符串，返回活动时间、加分时间和公示时间"""
    try:
        # 检查是否只有年月
        if re.match(r'^\d{4}\.\d{1,2}$', date_str):
            year = int(date_str.split('.')[0])
            month = int(date_str.split('.')[1])
            hold_date = f"{year:04d}.{month:02d}.15"
            bonus_date = f"{year:04d}.{month:02d}.16"
            pub_date = f"{year:04d}.{month:02d}.17"
        else:
            # 完整的日期格式
            match = re.match(r'(\d{4})\.(\d{1,2})\.(\d{1,2})', date_str)
            if match:
                year, month, day = map(int, match.groups())
                hold_date = f"{year:04d}.{month:02d}.{day:02d}"
                
                # 加分时间：活动时间+1天
                hold_dt = datetime.date(year, month, day)
                bonus_dt = hold_dt + datetime.timedelta(days=1)
                bonus_date = f"{bonus_dt.year:04d}.{bonus_dt.month:02d}.{bonus_dt.day:02d}"
                
                # 公示时间：活动时间+2天
                pub_dt = hold_dt + datetime.timedelta(days=2)
                pub_date = f"{pub_dt.year:04d}.{pub_dt.month:02d}.{pub_dt.day:02d}"
            else:
                raise ValueError(f"无法解析日期格式: {date_str}")
        
        return hold_date, bonus_date, pub_date
        
    except Exception as e:
        print(f"日期处理错误: {e}")
        # 返回默认日期
        default_date = "2025.01.01"
        return default_date, default_date, default_date

# --------------------------------------------------
# 活动数据处理类
# --------------------------------------------------

class ActivityProcessor:
    def __init__(self):
        self.activities = []
    
    def read_excel_files(self):
        """读取Input_输入目录下的所有Excel文件"""
        excel_files = list(INPUT_EXCEL_DIR.glob("*.xlsx"))
        
        if not excel_files:
            print(f"错误: 在 {INPUT_EXCEL_DIR} 目录下未找到Excel文件")
            return False
        
        all_data = []
        for excel_file in excel_files:
            try:
                print(f"正在读取: {excel_file}")
                df = pd.read_excel(excel_file)
                
                # 检查必要的列是否存在
                required_columns = ['班级', '姓名', '事由', '性质', '分数', '日期']
                missing_columns = [col for col in required_columns if col not in df.columns]
                if missing_columns:
                    print(f"警告: 文件 {excel_file} 缺少列: {missing_columns}")
                    continue
                
                # 添加文件名标识，便于追踪
                df['来源文件'] = excel_file.name
                all_data.append(df)
                
            except Exception as e:
                print(f"读取文件 {excel_file} 时出错: {e}")
        
        if not all_data:
            return False
        
        # 合并所有数据
        combined_df = pd.concat(all_data, ignore_index=True)
        
        # 清理数据
        combined_df = combined_df.dropna(subset=['班级', '姓名', '事由'])
        combined_df = combined_df.fillna('')
        
        # 处理事由字段
        combined_df['处理后事由'] = combined_df['事由'].apply(self.clean_activity_name)
        
        # 分组处理活动
        self.group_activities(combined_df)
        return True
    
    def clean_activity_name(self, activity_name):
        """清理活动名称"""
        activity_name = str(activity_name).strip()
        # 去掉开头的"参加"、"参与"
        if activity_name.startswith('参加'):
            activity_name = activity_name[2:].strip()
        elif activity_name.startswith('参与'):
            activity_name = activity_name[2:].strip()
        return activity_name
    
    def group_activities(self, df):
        """按事由、性质、分数分组活动数据"""
        # 首先按事由、性质、分数分组
        grouped = df.groupby(['处理后事由', '性质', '分数'])
        
        for (activity, score_type, score), group in grouped:
            # 处理分数格式
            score_value = self.extract_score(score)
            
            # 获取日期（取组内第一个非空日期）
            date_str = str(group['日期'].iloc[0]) if not group['日期'].isna().all() else "2025.01.01"
            
            
            # 创建活动对象
            activity_data = {
                'activity_name': activity,
                'score_type': score_type,
                'score_value': score_value,
                'date_str': date_str,
                'participants': []
            }
            
            # 添加参与者
            for _, row in group.iterrows():
                participant = {
                    'name': str(row['姓名']).strip(),
                    'class_name': convert_class_name(str(row['班级']).strip())
                }
                activity_data['participants'].append(participant)
            
            self.activities.append(activity_data)
        
        print(f"共找到 {len(self.activities)} 个不同的活动\n")
    
    def extract_score(self, score):
        """从分数字符串中提取数值"""
        try:
            score_str = str(score).strip()
            # 去掉+号和其他非数字字符
            score_clean = re.sub(r'[^\d]', '', score_str)
            return int(score_clean) if score_clean else 0
        except:
            return 0

# --------------------------------------------------
# 数据文件生成器
# --------------------------------------------------

class DataFileGenerator:
    def __init__(self, activity_processor):
        self.processor = activity_processor
        self.template_content = ""
        self.output_lines = []
    
    def load_template(self):
        """加载模板文件"""
        try:
            with open(TEMPLATE_DATA_FILE, 'r', encoding='utf-8') as f:
                self.template_content = f.read()
            return True
        except Exception as e:
            print(f"加载模板文件失败: {e}")
            return False
    
    def generate_data_file(self):
        """生成完整的数据文件"""
        if not self.load_template():
            return False
        
        # 解析模板，找到需要替换的部分
        lines = self.template_content.split('\n')
        output_lines = []
        
        current_section = None
        in_field_block = False
        current_field = None
        
        for line in lines:
            stripped = line.strip()
            
            # 检查是否是字段定义行
            if stripped.startswith("-  $ "):
                field_name = self.extract_field_name(stripped)
                current_field = field_name
                in_field_block = True
                output_lines.append(line)
            
            # 检查是否是占位符行
            elif stripped.startswith("-  @ ") and in_field_block:
                field_name = self.extract_field_name(stripped)
                if field_name == current_field:
                    # 替换这个占位符
                    self.replace_field_data(output_lines, current_field)
                else:
                    output_lines.append(line)
            
            else:
                output_lines.append(line)
                if not stripped and in_field_block:
                    in_field_block = False
                    current_field = None
        
        # 写入输出文件
        try:
            with open(OUTPUT_DATA_FILE, 'w', encoding='utf-8') as f:
                f.write('\n'.join(output_lines))
            print(f"✓ 已生成数据文件: {OUTPUT_DATA_FILE}\n")
            return True
        except Exception as e:
            print(f"✗ 写入数据文件失败: {e}")
            return False
    
    def extract_field_name(self, line):
        """从字段行中提取字段名"""
        # 去掉 "-  $ " 或 "-  @ " 前缀
        field_line = line[5:]
        # 提取字段名（去掉注释部分）
        field_name = field_line.split(':', 1)[0].strip()
        # 去掉方括号内容
        field_name = re.sub(r'\[.*?\]', '', field_name).strip()
        return field_name
    
    def replace_field_data(self, output_lines, field_name):
        """替换特定字段的数据"""
        if field_name == "Output.files":
            self.generate_output_files(output_lines)
        elif field_name == "RN.split(，)":
            self.generate_rn_split(output_lines)
        elif field_name == "RC.split(，)":
            self.generate_rc_split(output_lines)
        elif field_name == "Reason.Activity":
            self.generate_reason_activity(output_lines)
        elif field_name == "Reason.Page.int":
            self.generate_reason_page(output_lines)
        elif field_name == "Score.type":
            self.generate_score_type(output_lines)
        elif field_name == "NoP.CL.int":
            self.generate_nop(output_lines)
        elif field_name == "Score.CL.int":
            self.generate_score_value(output_lines)
        elif field_name in ["Time.Act.Hold.dateD", "Time.Act.Bonus.dateD", "Time.Act.Pub.dateD", "Time.Act.Pub.date"]:
            self.generate_time_fields(output_lines, field_name)
        else:
            # 对于其他字段，保持原样
            output_lines.append(field_name)
    
    def generate_output_files(self, output_lines):
        """生成Output.files数据"""
        file_pairs = []
        
        for i, activity in enumerate(self.processor.activities):
            activity_name = activity['activity_name']
            participants = activity['participants']
            
            # 计算需要的页数（每页最多20人）
            num_pages = (len(participants) + 19) // 20
            
            for page in range(num_pages):
                suffix = f"_{i+1}" if len(self.processor.activities) > 1 else ""
                if num_pages > 1:
                    suffix += f"_{page+1}"
                
                approval_file = f"院审批单_{activity_name}{suffix}.docx"
                receipt_file = f"院回执单_{activity_name}{suffix}.docx"
                file_pairs.append(f"{approval_file} {receipt_file}")
        
        # 添加到输出
        for file_pair in file_pairs:
            output_lines.append(file_pair)
    
    def generate_rn_split(self, output_lines):
        """生成RN.split数据"""
        for activity in self.processor.activities:
            participants = activity['participants']
            
            # 分页处理（每页最多20人）
            num_pages = (len(participants) + 19) // 20
            
            for page in range(num_pages):
                start_idx = page * 20
                end_idx = min((page + 1) * 20, len(participants))
                page_participants = participants[start_idx:end_idx]
                
                names = [p['name'] for p in page_participants]
                names_str = "，".join(names)
                output_lines.append(names_str)
    
    def generate_rc_split(self, output_lines):
        """生成RC.split数据"""
        for activity in self.processor.activities:
            participants = activity['participants']
            
            # 分页处理（每页最多20人）
            num_pages = (len(participants) + 19) // 20
            
            for page in range(num_pages):
                start_idx = page * 20
                end_idx = min((page + 1) * 20, len(participants))
                page_participants = participants[start_idx:end_idx]
                
                classes = [p['class_name'] for p in page_participants]
                classes_str = "，".join(classes)
                output_lines.append(classes_str)
    
    def generate_reason_activity(self, output_lines):
        """生成Reason.Activity数据"""
        for activity in self.processor.activities:
            activity_name = activity['activity_name']
            participants = activity['participants']
            
            # 每个活动重复的次数等于页数
            num_pages = (len(participants) + 19) // 20
            for _ in range(num_pages):
                output_lines.append(activity_name)
    
    def generate_reason_page(self, output_lines):
        """生成Reason.Page.int数据"""
        for activity in self.processor.activities:
            participants = activity['participants']
            num_pages = (len(participants) + 19) // 20
        
            # 如果只需要1页，所有输出都是页码1
            if num_pages == 1:
                output_lines.append("1")
            else:
                # 需要多页时，按分页分配页码
                for page in range(num_pages):
                    output_lines.append(str(page + 1))
    
    def generate_score_type(self, output_lines):
        """生成Score.type数据"""
        for activity in self.processor.activities:
            score_type = activity['score_type']
            participants = activity['participants']
            
            num_pages = (len(participants) + 19) // 20
            for _ in range(num_pages):
                output_lines.append(score_type)
    
    def generate_nop(self, output_lines):
        """生成NoP.CL.int数据"""
        for activity in self.processor.activities:
            participants = activity['participants']
            
            num_pages = (len(participants) + 19) // 20
            for page in range(num_pages):
                start_idx = page * 20
                end_idx = min((page + 1) * 20, len(participants))
                page_count = end_idx - start_idx
                
                output_lines.append(str(page_count))
    
    def generate_score_value(self, output_lines):
        """生成Score.CL.int数据"""
        for activity in self.processor.activities:
            score_value = activity['score_value']
            participants = activity['participants']
            
            num_pages = (len(participants) + 19) // 20
            for _ in range(num_pages):
                output_lines.append(str(score_value))
    
    def generate_time_fields(self, output_lines, field_name):
        """生成时间字段数据"""
        time_type_map = {
            "Time.Act.Hold.dateD": 0,  # 活动时间
            "Time.Act.Bonus.dateD": 1,  # 加分时间
            "Time.Act.Pub.dateD": 2,  # 公示时间
            "Time.Act.Pub.date": 1  # 审批单上的时间
        }
        
        time_type = time_type_map.get(field_name, 0)
        
        for activity in self.processor.activities:
            date_str = activity['date_str']
            hold_date, bonus_date, pub_date = process_date(date_str)
            
            dates = [hold_date, bonus_date, pub_date]
            target_date = dates[time_type]
            
            participants = activity['participants']
            num_pages = (len(participants) + 19) // 20
            
            for _ in range(num_pages):
                output_lines.append(target_date)

# --------------------------------------------------
# 主程序
# --------------------------------------------------

def main():
    # 检查必要的目录和文件
    if not INPUT_EXCEL_DIR.exists():
        INPUT_EXCEL_DIR.mkdir(parents=True, exist_ok=True)
        print(f"已创建输入目录: {INPUT_EXCEL_DIR}")
        print("请将月报表Excel文件放入该目录后重新运行程序")
        return
    
    if not TEMPLATE_DATA_FILE.exists():
        print(f"错误: 模板文件不存在 {TEMPLATE_DATA_FILE}")
        print("请确保在 Template_模板 目录下创建 DR_data_Template.txt 文件")
        return
    
    # 处理月报表数据
    print("步骤1: 读取月报表数据...")
    processor = ActivityProcessor()
    if not processor.read_excel_files():
        print("未能读取到有效的月报表数据")
        return
    
    if not processor.activities:
        print("未找到有效的活动数据")
        return
    
    # 生成数据文件
    print("步骤2: 生成数据文件...")
    generator = DataFileGenerator(processor)
    if not generator.generate_data_file():
        print("生成数据文件失败")
        return
    
    # 执行文档替换脚本
    print("步骤3: 调用文档替换函数...\n")
    if SCRIPT_FILE_1.exists() and SCRIPT_FILE_2.exists():
        import DocxReplacer
        DocxReplacer.main()
        import Combiner
        Combiner.main()
    else:
        print(f"错误: 脚本文件不存在 {SCRIPT_FILE_1} {SCRIPT_FILE_2}")
        return

if __name__ == "__main__":
    print('='*25)
    print(f"正在运行 {os.path.splitext(os.path.basename(__file__))[0]} ...")
    print('='*25, "\n")
    main()
    input("\n按回车键退出...\n")
else:
    print('='*25)
    print(f"正在加载 {__name__} ...")
    print('='*25, "\n")