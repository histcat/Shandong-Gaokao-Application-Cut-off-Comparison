#!/usr/bin/env python3
"""
志愿-投档线对照分析脚本

对于志愿表中的每一个志愿，按照报考院校代码和专业代码查询投档线中的最低位次，
与用户输入的位次比较，判断是否会被录取。
当遇到会被录取的志愿时，输出结果并停止。

用法: python compare.py
"""

import pandas as pd
import PyPDF2
import re
import sys
import os
import unicodedata

# -- 编码处理 ------------------------------------------
# Windows GBK终端可能无法输出某些CJK兼容字符，设置UTF-8并做兼容处理
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass


def sanitize(text):
    """将CJK兼容字符等转为普通中文，避免GBK编码错误"""
    # 先将CJK兼容字符转为标准形式
    text = unicodedata.normalize('NFKC', text)
    # 用 ? 替换无法编码的字符
    result = []
    for ch in text:
        try:
            ch.encode('gbk')
            result.append(ch)
        except UnicodeEncodeError:
            # 尝试用标准等价字符
            try:
                decomposed = unicodedata.normalize('NFKD', ch)
                # 取第一个可编码的基本字符
                for dc in decomposed:
                    try:
                        dc.encode('gbk')
                        result.append(dc)
                        break
                    except UnicodeEncodeError:
                        continue
                else:
                    result.append('?')
            except Exception:
                result.append('?')
    return ''.join(result)


# -- 配置 ----------------------------------------------
PDF_PATH = "志愿表.pdf"
XLS_PATH = "投档线.xls"


# -- PDF 解析 ------------------------------------------

def extract_pdf_text(pdf_path):
    """从PDF中提取全部文本"""
    reader = PyPDF2.PdfReader(pdf_path)
    full_text = ""
    for page in reader.pages:
        text = page.extract_text()
        if text:
            full_text += text + "\n"
    return full_text


def parse_volunteer_table(text):
    """
    解析志愿表文本，返回志愿列表。
    每条志愿: {志愿号, 院校代码, 院校名称, 专业代码, 专业名称, 层次, 年收费}
    """
    lines = text.split('\n')

    # 找到表头行（包含"志愿号"的行），从它之后开始解析
    start_idx = 0
    for i, line in enumerate(lines):
        if '志愿号' in line and '报考院校' in line:
            start_idx = i + 1
            break

    # 志愿起始行模式: 行首为数字 + 可选空格 + 大写字母+3位数字(院校代码)
    # 注意：有些行如 "53A358中国科学技术大" 没有空格分隔
    entry_pattern = re.compile(r'^(\d+)\s*([A-Z]\d{3})')

    # 按条目收集行
    raw_entries = []
    current_lines = []

    for i in range(start_idx, len(lines)):
        line = lines[i]
        match = entry_pattern.match(line)
        if match:
            if current_lines:
                raw_entries.append(''.join(current_lines))
            current_lines = [line]
        else:
            if current_lines:  # 续行
                current_lines.append(line)

    if current_lines:
        raw_entries.append(''.join(current_lines))

    # 解析每个条目
    volunteers = []
    for entry_text in raw_entries:
        vol = parse_single_entry(entry_text)
        if vol:
            volunteers.append(vol)

    return volunteers


def parse_single_entry(text):
    """解析单条志愿的文本"""
    if '查看时间' in text:
        text = text.split('查看时间')[0]

    text = text.strip()
    if not text:
        return None

    # 提取志愿号: 行首数字
    vol_match = re.match(r'^(\d+)\s*', text)
    if not vol_match:
        return None
    vol_num = int(vol_match.group(1))

    # 提取院校代码: [A-Z]后跟3位数字
    school_code_match = re.search(r'([A-Z]\d{3})', text)
    if not school_code_match:
        return None
    school_code = school_code_match.group(1)

    # 提取院校名称: 院校代码之后到办学性质之前
    school_start = school_code_match.end()
    nature_pattern = re.search(r'(公办院校|民办院校|独立学院|中外合作办学|内地与港澳台地区合作办学)', text)
    if nature_pattern:
        school_name = text[school_start:nature_pattern.start()].strip()
        after_nature = text[nature_pattern.end():]
    else:
        school_name = text[school_start:].strip()
        after_nature = text

    # 从办学性质后提取专业代码和名称
    major_code_match = re.match(r'\s*([0-9A-Za-z]+)\s*', after_nature)
    if not major_code_match:
        return None
    major_code = major_code_match.group(1)
    rest = after_nature[major_code_match.end():].strip()

    # 提取专业名称（到"本科"/"专科"之前）
    level_match = re.search(r'(本科|专科|高职)', rest)
    if level_match:
        full_major_text = rest[:level_match.start()].strip()
        level = level_match.group(1)
        after_level = rest[level_match.end():]
    else:
        full_major_text = rest.strip()
        level = ""
        after_level = ""

    # 提取年收费
    fee_match = re.search(r'(\d{3,6})', after_level) if after_level else None
    fee = fee_match.group(1) if fee_match else ""

    # 简化专业名称用于显示
    short_name = re.split(r'[（(]', full_major_text)[0].strip()
    short_name = re.sub(r'\s+', '', short_name)

    return {
        '志愿号': vol_num,
        '院校代码': school_code,
        '院校名称': sanitize(school_name),
        '专业代码': major_code,
        '专业名称': sanitize(short_name),
        '专业全称': sanitize(full_major_text),
        '层次': level,
        '年收费': fee,
    }


# -- 投档线XLS解析 -------------------------------------

def parse_toudang_xls(xls_path):
    """
    解析投档线Excel，返回:
      lookup: {(院校代码, 专业代码): {'院校名称', '专业名称', '投档计划数', '最低位次'}}
    """
    df = pd.read_excel(xls_path, header=None, skiprows=1)
    df.columns = ['专业代码及名称', '院校代码及名称', '投档计划数', '最低位次']

    lookup = {}

    for _, row in df.iterrows():
        major_full = str(row['专业代码及名称']).strip()
        school_full = str(row['院校代码及名称']).strip()
        plan = str(row['投档计划数']).strip()
        min_rank_raw = str(row['最低位次']).strip()

        # 提取院校代码
        school_match = re.match(r'^([A-Z]\d{3})', school_full)
        if not school_match:
            continue
        school_code = school_match.group(1)
        school_name = school_full[school_match.end():]

        # 提取专业代码
        major_match = re.match(r'^([0-9A-Za-z]+)', major_full)
        if not major_match:
            continue
        major_code = major_match.group(1)
        major_name = major_full[major_match.end():]

        # 解析最低位次
        try:
            min_rank = int(min_rank_raw)
        except ValueError:
            # "前50名"等特殊情况，提取数字
            rank_clean = re.sub(r'[^\d]', '', min_rank_raw)
            min_rank = int(rank_clean) if rank_clean else None

        lookup[(school_code, major_code)] = {
            '院校名称': sanitize(school_name),
            '专业名称': sanitize(major_name),
            '投档计划数': plan,
            '最低位次': min_rank,
            '最低位次原始': min_rank_raw,
        }

    return lookup


# -- 主逻辑 --------------------------------------------

def main():
    if not os.path.exists(PDF_PATH):
        print(f"[ERROR] 找不到文件: {PDF_PATH}")
        sys.exit(1)
    if not os.path.exists(XLS_PATH):
        print(f"[ERROR] 找不到文件: {XLS_PATH}")
        sys.exit(1)

    print("=" * 70)
    print("           志愿-投档线对照分析")
    print("=" * 70)

    # 解析PDF
    print("\n[PDF] 正在解析志愿表...")
    text = extract_pdf_text(PDF_PATH)

    info_match = re.search(r'姓名[：:]\s*(\S+)', text)
    score_match = re.search(r'成绩[：:]\s*(\d+)\s*(\d+)', text)
    student_name = info_match.group(1) if info_match else "未知"
    student_score = f"{score_match.group(1)}{score_match.group(2)}" if score_match else "未知"
    print(f"   考生: {sanitize(student_name)}  成绩: {student_score}")

    volunteers = parse_volunteer_table(text)
    print(f"   共解析到 {len(volunteers)} 条志愿")

    # 解析投档线
    print("[XLS] 正在解析投档线...")
    lookup = parse_toudang_xls(XLS_PATH)
    print(f"   共解析到 {len(lookup)} 条投档记录")

    # 获取用户位次
    print("\n" + "-" * 70)
    while True:
        try:
            user_input = input("请输入你的位次: ").strip()
            user_rank = int(user_input)
            if user_rank <= 0:
                print("位次必须为正整数，请重新输入。")
                continue
            break
        except ValueError:
            print("请输入有效的整数位次。")
        except (EOFError, KeyboardInterrupt):
            print("\n程序终止。")
            sys.exit(0)

    print(f"你的位次: {user_rank}")
    print("-" * 70)

    # 逐条比对
    admitted = False
    not_found_count = 0  # 记录未找到的数量

    for vol in volunteers:
        vol_num = vol['志愿号']
        school_code = vol['院校代码']
        school_name = vol['院校名称']
        major_code = vol['专业代码']
        major_name = vol['专业名称']

        print(f"\n{'-' * 60}")
        print(f"志愿 {vol_num}: [{school_code}] {school_name}")
        print(f"  专业代码: {major_code}  {major_name}")

        # 精确代码匹配
        key = (school_code, major_code)
        if key not in lookup:
            print(f"  [WARN]  投档线中未找到该代码组合，无法判断")
            not_found_count += 1
            continue

        info = lookup[key]
        min_rank = info['最低位次']

        print(f"  投档专业: {info['专业名称']}")
        print(f"  投档计划数: {info['投档计划数']}")
        print(f"  最低位次: {info['最低位次原始']}")

        if min_rank is None:
            print(f"  [WARN]  最低位次无法解析，跳过")
            continue

        # 判断录取: 用户位次 <= 最低位次 → 录取（位次越小排名越靠前）
        if user_rank <= min_rank:
            print(f"  [OK] 录取！你的位次({user_rank}) ≤ 最低位次({min_rank})")
            admitted = True
            print(f"\n{'=' * 70}")
            print(f"  *** 在第 {vol_num} 志愿被录取！")
            print(f"     院校: [{school_code}] {school_name}")
            print(f"     专业: [{major_code}] {major_name}")
            print(f"     你的位次: {user_rank}  ≤  最低位次: {min_rank}")
            print(f"{'=' * 70}")
            break
        else:
            print(f"  [NO] 未录取。你的位次({user_rank}) > 最低位次({min_rank})")

    if not admitted:
        print(f"\n{'=' * 70}")
        if not_found_count == len(volunteers):
            print(f"  --- 所有 {len(volunteers)} 个志愿在投档线中均未找到匹配记录。")
        else:
            print(f"  --- 遍历全部 {len(volunteers)} 个志愿，均未达到投档线。")
        print(f"{'=' * 70}")

    print()


if __name__ == '__main__':
    main()
