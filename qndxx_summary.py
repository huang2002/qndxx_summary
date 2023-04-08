#!/usr/bin/env python3
# -*- coding: utf-8 -*-

##################################################
# 青年大学习统计程序
# https://github.com/huang2002/qndxx_summary
##################################################

import os
import re
import time
from typing import Any, Callable, Dict, List, NoReturn, NamedTuple


def halt(code: object) -> NoReturn:
    print('程序终止。\n')
    input('请按回车键退出……')
    exit(code)


try:
    import pandas as pd
except ImportError:
    print('此程序需要用到pandas，请先使用以下命令安装：')
    print('pip install pandas')
    print()
    halt(1)

### Initialization. ###

RECORD_DIR_NAME = 'records'  # 观看记录文件夹名称
STUDENT_DIR_NAME = 'students'  # 团员列表文件夹名称
OUTPUT_DIR_NAME = 'output'  # 输出文件夹名称

BASE_DIR: str = os.path.dirname(__file__)
RECORD_DIR: str = os.path.join(BASE_DIR, RECORD_DIR_NAME)
STUDENT_DIR: str = os.path.join(BASE_DIR, STUDENT_DIR_NAME)
OUTPUT_BASE_DIR: str = os.path.join(BASE_DIR, OUTPUT_DIR_NAME)

################################################
# 这些变量指定了输入文件的编码格式。
# 当程序报错提示与编码有关时，可以尝试修改这个变量。
# （Excel常用“gbk”和“gb2312”，CSV常用“utf-8”。）
################################################
RECORD_ENCODING = 'gbk'  # 观看记录文件编码
STUDENT_ENCODING = 'gb2312'  # 团员列表文件编码

#######################################################
# 这些是观看记录文件表头名称相关的变量。
# KEY_RECORD_*是后面程序会直接用到的表头；
# RECORD_COLUMNS是完整的表头列表，
# 必须和观看记录文件中的列对应。
# （程序将无视观看记录文件的表头，而是根据这里的声明处理数据。）
#######################################################
KEY_RECORD_ISSUE = '课程'
KEY_RECORD_IDENTITY = '学号/卡号/工号'
KEY_RECORD_CLASS = '班级'
KEY_RECORD_TIME = '学习时间'
KEY_RECORD_NAME = '姓名'
RECORD_COLUMNS = [
    KEY_RECORD_ISSUE,
    '系统',
    '学校',
    '学院',
    KEY_RECORD_CLASS,
    KEY_RECORD_IDENTITY,
    KEY_RECORD_TIME,
]

############################################
# 这些是团员列表文件表头名称相关的变量。
# KEY_STUDENT_*是后面程序会直接用到的表头；
# STUDENT_COLUMNS是将要使用的表头列表，
# 必须存在于student文件夹里的团员列表文件中。
############################################
KEY_STUDENT_CLASS = '班级'
KEY_STUDENT_NAME = '姓名'
KEY_STUDENT_ID = '学号'
STUDENT_COLUMNS = [
    KEY_STUDENT_CLASS,
    KEY_STUDENT_NAME,
    KEY_STUDENT_ID,
]

####################################################
# 输出的符号，positive指“已观看”，negative指“未观看”。
####################################################
OUTPUT_POSITIVE = '√'
OUTPUT_NEGATIVE = '×'

common_reader_params = {
    'skiprows': 1,
    'skipfooter': 1,
    'parse_dates': [KEY_RECORD_TIME],
    'infer_datetime_format': True,
}

FileReader = Callable[[str], pd.DataFrame]  # path -> df_records


def read_csv(path: str, **addition_params: Dict[str, Any]) -> pd.DataFrame:
    return pd.read_csv(
        path,
        engine='python',
        **common_reader_params,
        **addition_params,
    )


def read_excel(path: str, **addition_params: Dict[str, Any]) -> pd.DataFrame:
    return pd.read_excel(
        path,
        **common_reader_params,
        **addition_params,
    )


# lowercased_file_extension -> reader
file_readers: Dict[str, FileReader] = {
    '.csv': read_csv,
    '.xlsx': read_excel,
}


def read_file(path: str, **addition_params: Dict[str, Any]) -> pd.DataFrame:
    file_extension = os.path.splitext(filename)[1]
    file_reader = file_readers[file_extension.lower()]
    return file_reader(file_path, **addition_params)


def is_acceptable_filename(filename: str) -> bool:
    file_extension: str = os.path.splitext(filename)[1]
    return file_extension.lower() in file_readers


def list_acceptable_filenames(directory_path: str) -> List[str]:
    return list(
        filter(
            is_acceptable_filename,
            os.listdir(directory_path)
        )
    )


StudentIdentity = NamedTuple('StudentIdentity', id=str, name=str)

# match_result -> student_identity
IdentityParser = Callable[[re.Match], StudentIdentity]

# pattern -> parser
identity_parsers = {
    # id, name
    re.compile(r'^([a-zA-Z\d]+)([^a-zA-Z\d]+)$'): (lambda m: (m[1], m[2])),
    # name, id
    re.compile(r'^([^a-zA-Z\d]+)([a-zA-Z\d]+)$'): (lambda m: (m[2], m[1])),
    # id only
    re.compile(r'^([a-zA-Z\d]+)$'): (lambda m: (m[1], '')),
    # name only
    re.compile(r'^([^a-zA-Z\d]+)$'): (lambda m: ('', m[1])),
}

### Start up. ###

print('青年大学习统计程序 v0.1.0')
print()

print('当前支持的文件后缀：（不区分大小写）')
for ext in file_readers.keys():
    print(f'  {ext}')
print()

### Check directories. ###

if not os.path.exists(RECORD_DIR):
    os.mkdir(RECORD_DIR)
    print('已自动创建观看记录文件夹。')
    print()
    input(f'请将所有观看记录文件放入{RECORD_DIR_NAME}文件夹，然后按回车键继续……')
    print()

if not os.path.exists(STUDENT_DIR):
    os.mkdir(STUDENT_DIR)
    print('已自动创建团员列表文件夹。')
    print()
    input(f'请将团员列表文件放入{STUDENT_DIR_NAME}文件夹，然后按回车键继续……')
    print()

### Check record files. ###

record_filenames: List[str] = list_acceptable_filenames(RECORD_DIR)
if len(record_filenames) == 0:
    print('未检测到观看记录文件！')
    halt(1)

print('检测到观看记录文件：')
for filename in record_filenames:
    print(f'  {filename}')
print()

### Check student files. ###

student_filenames: List[str] = list_acceptable_filenames(STUDENT_DIR)
if len(student_filenames) == 0:
    print('未检测到团员名单文件。')
else:
    print('检测到团员名单文件：')
    for filename in student_filenames:
        print(f'  {filename}')
print()

### Read record files. ###

print('加载观看记录文件：')
record_dataframes: List[pd.DataFrame] = []
for filename in record_filenames:
    print(f'  读入{filename}……')
    file_path = os.path.join(RECORD_DIR, filename)
    df_record = read_file(file_path, encoding=RECORD_ENCODING)
    df_record.columns = RECORD_COLUMNS
    df_record.dropna(inplace=True)
    df_record.drop_duplicates(inplace=True)
    record_dataframes.append(df_record)
print()

### Load student data. ###

print('加载团员列表……')

StudentInfo = NamedTuple('StudentInfo', class_=str, name=str)
students: List[StudentInfo] = []

student_name_map = {}  # id -> name
for filename in student_filenames:
    print(f'  读入{filename}……')
    file_path = os.path.join(STUDENT_DIR, filename)
    df_students = read_file(file_path, encoding=STUDENT_ENCODING)
    df_students.columns = STUDENT_COLUMNS
    df_students.dropna(inplace=True)
    df_students.drop_duplicates(inplace=True)
    for index in df_students.index:
        student_id = df_students.at[index, KEY_STUDENT_ID]
        if student_id in student_name_map:
            print(f'  发现重复的学号：{student_id}')
            halt(1)
        student_name = df_students.at[index, KEY_STUDENT_NAME]
        student_class = df_students.at[index, KEY_STUDENT_CLASS]
        student_name_map[student_id] = student_name
        student_info = StudentInfo(class_=student_class, name=student_name)
        students.append(student_info)

print()

### Process records. ###

print('处理观看记录……')


def identity_to_name(identity: str, filename: str) -> str:
    for pattern, parser in identity_parsers.items():
        match_result = pattern.match(identity)
        if not match_result:
            continue
        student_id, student_name = parser(match_result)
        if student_name:
            return student_name
        elif student_id in student_name_map:
            return student_name_map[student_id]
        else:
            return student_id


IssueInfo = NamedTuple('IssueInfo', name=str, time=pd.DatetimeTZDtype)
issues: List[IssueInfo] = []

for i, df_records in enumerate(record_dataframes):

    df_issues = df_records[[KEY_RECORD_ISSUE, KEY_RECORD_TIME]] \
        .groupby(KEY_RECORD_ISSUE) \
        .mean()
    for issue_name in df_issues.index:
        if any((issue.name == issue_name) for issue in issues):
            continue
        issue_time = df_issues.at[issue_name, KEY_RECORD_TIME]
        issue = IssueInfo(name=issue_name, time=issue_time)
        issues.append(issue)

    filename = record_filenames[i]
    df_records[KEY_RECORD_NAME] = df_records[KEY_RECORD_IDENTITY].map(
        lambda identity: identity_to_name(identity, filename)
    )
    for index in df_records.index:
        student_class = df_records.at[index, KEY_RECORD_CLASS]
        student_name = df_records.at[index, KEY_RECORD_NAME]
        student_info = StudentInfo(class_=student_class, name=student_name)
        if not student_info in students:
            students.append(student_info)

# Sort issues in ascending order.
issues: List[IssueInfo] = sorted(issues, key=(lambda issue: issue.time))

issue_names: List[str] = [issue.name for issue in issues]
student_classes, student_names = zip(*students)

print()

### Generate output data. ###

print('生成统计结果……')

init_data = [
    dict(
        (issue.name, OUTPUT_NEGATIVE)
        for issue in issues
    )
    for student in students
]
init_index = pd.MultiIndex.from_tuples(
    [(student.class_, student.name) for student in students],
    names=[KEY_RECORD_CLASS, KEY_RECORD_NAME],
)

df_output = pd.DataFrame(
    data=init_data,
    index=init_index,
)

for df_records in record_dataframes:
    for index in df_records.index:
        student_class = df_records.at[index, KEY_RECORD_CLASS]
        student_name = df_records.at[index, KEY_RECORD_NAME]
        issue_name = df_records.at[index, KEY_RECORD_ISSUE]
        student_index = (student_class, student_name)
        df_output.at[student_index, issue_name] = OUTPUT_POSITIVE

df_output.reset_index(inplace=True)
df_output.sort_values(by=KEY_RECORD_CLASS, inplace=True)

print()

### Generate result files. ###

print('输出结果……')

if not os.path.exists(OUTPUT_BASE_DIR):
    os.mkdir(OUTPUT_BASE_DIR)

timestamp: str = time.strftime('%Y-%m-%d_%H.%M.%S')
output_directory: str = os.path.join(OUTPUT_BASE_DIR, timestamp)
if os.path.exists(output_directory):
    print('  目标文件夹已存在，请稍后重试。')
    halt(1)
os.mkdir(output_directory)


def output_slice(indexer: slice, class_name: str) -> NoReturn:
    print(f'  导出{class_name}……')
    filename = class_name + '.xlsx'
    output_path = os.path.join(output_directory, filename)
    df_output.loc[indexer, :].to_excel(output_path, index=False)


begin_index = df_output.index[0]
current_class = df_output.at[begin_index, KEY_RECORD_CLASS]

for index in df_output.index:

    student_class = df_output.at[index, KEY_RECORD_CLASS]
    if student_class == current_class:
        continue

    output_slice(slice(begin_index, index), current_class)

    begin_index = index
    current_class = student_class

# last class
output_slice(slice(begin_index, None), current_class)

print()

### The end. ###

input('程序成功结束，请按回车键退出……')
