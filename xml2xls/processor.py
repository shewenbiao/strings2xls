import os
import re
import argparse
from openpyxl import Workbook, load_workbook


def export_to_excel(res_dir, output_file):
    """增强版导出功能，包含未翻译统计"""
    try:
        # 获取默认语言内容
        default_order, default_data = parse_strings_xml(os.path.join(res_dir, 'values', 'strings.xml'))
        all_langs = {'en': default_data}

        # 收集其他语言数据
        lang_codes = []
        for d in os.listdir(res_dir):
            if d.startswith('values-'):
                lang_code = d.replace('values-', '', 1)
                _, lang_data = parse_strings_xml(os.path.join(res_dir, d, 'strings.xml'))
                all_langs[lang_code] = lang_data
                lang_codes.append(lang_code)

        # 创建Excel文件
        wb = Workbook()

        # Sheet1: 完整翻译表
        main_sheet = wb.active
        main_sheet.title = "All Translations"
        main_sheet.append(['key', 'en'] + lang_codes)

        # 填充主表数据
        for key in default_order:
            row = [key, default_data.get(key, '')]
            for lc in lang_codes:
                row.append(all_langs[lc].get(key, ''))
            main_sheet.append(row)

        # Sheet2+: 各语言未翻译项
        for lang in lang_codes:
            untranslated = {}
            for key in default_data:
                if key not in all_langs[lang]:
                    untranslated[key] = default_data[key]

            if untranslated:
                sheet = wb.create_sheet(title=lang[:31])  # Excel表名最长31字符
                sheet.append(['key', 'en'])
                for key, value in untranslated.items():
                    sheet.append([key, value])

        wb.save(output_file)
        print(f"导出成功：{output_file}")

    except Exception as e:
        print(f"导出失败：{str(e)}")
        raise


def import_from_excel(res_dir, input_file):
    """智能导入，保留原始格式"""
    try:
        wb = load_workbook(input_file)
        ws = wb.active

        # 解析表头
        headers = [cell.value for cell in ws[1]]
        lang_codes = headers[1:]  # 排除key列

        lang_data = {lc: {} for lc in lang_codes}
        for row in ws.iter_rows(min_row=2):
            key = row[0].value
            if not key:
                continue

            for idx, lc in enumerate(lang_codes, start=1):
                value = row[idx].value
                if value:
                    lang_data[lc][key] = str(value).strip()

        # 写入各语言文件
        for lang_code, data in lang_data.items():
            dir_name = 'values' if lang_code == 'en' else f'values-{lang_code}'
            xml_path = os.path.join(res_dir, dir_name, 'strings.xml')
            write_strings_xml(xml_path, data)

        print(f"导入成功！")
    except Exception as e:
        print(f"导入失败：{str(e)}")
        raise


def parse_strings_xml(xml_path):
    """解析XML获取原始内容（保留CDATA等特殊格式）"""
    order = []
    strings = {}
    if not os.path.exists(xml_path):
        return order, strings

    with open(xml_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # 增强型正则表达式，匹配完整字符串定义
    pattern = re.compile(
        r'<!--(.*?)-->|'  # 匹配注释
        r'<string\s+name="([^"]+)"\s*>(.*?)</string>|'
        r'(</?resources>)',  # 匹配根标签
        re.DOTALL
    )

    for match in re.finditer(pattern, content):
        if match.group(1):  # 注释
            continue
        elif match.group(2):  # string标签
            name = match.group(2)
            value = match.group(3).strip()
            order.append(name)
            strings[name] = value
        elif match.group(4):  # resources标签
            continue  # 根标签不处理

    return order, strings


def write_strings_xml(xml_path, data):
    """智能合并写入XML（保留注释等其他内容）"""
    # 读取原始文件内容
    if os.path.exists(xml_path):
        with open(xml_path, 'r', encoding='utf-8') as f:
            content = f.read()
    else:
        content = '<?xml version="1.0" encoding="utf-8"?>\n<resources>\n</resources>'

    # 获取现有字符串和顺序
    existing_order, existing_data = parse_strings_xml(xml_path)

    # 合并数据：更新现有值，保留未修改内容
    merged_data = existing_data.copy()
    merged_data.update(data)
    new_order = existing_order + [k for k in data if k not in existing_order]

    # 构建替换字典（包含原始格式）
    replace_dict = {}
    for match in re.finditer(r'<string\s+name="([^"]+)"\s*>(.*?)</string>', content, re.DOTALL):
        name = match.group(1)
        if name in merged_data:
            old_content = match.group(0)
            new_content = f'<string name="{name}">{merged_data[name]}</string>'
            replace_dict[old_content] = new_content

    # 执行替换
    for old, new in replace_dict.items():
        content = content.replace(old, new, 1)

    # 添加新条目（追加到文件末尾前）
    new_entries = []
    for name in new_order:
        if name not in existing_data:
            new_entries.append(f'    <string name="{name}">{data[name]}</string>')

    if new_entries:
        # 在最后一个</string>后或</resources>前插入
        # 查找最后一个</string>的结束位置
        last_string_end = 0
        last_string_pos = content.rfind('</string>')
        if last_string_pos != -1:
            last_string_end = last_string_pos + len('</string>')

        # 查找</resources>的位置
        resources_end_pos = content.find('</resources>')

        # 确定插入位置（优先插入在最后一个</string>之后）
        insert_pos = last_string_end if last_string_end > 0 else resources_end_pos

        # 构建插入文本（包含换行格式）
        insert_text = '\n' + '\n'.join(f'{entry}' for entry in new_entries)

        # 执行插入（在插入位置后添加内容）
        content = content[:insert_pos] + insert_text + content[insert_pos:]

    # 处理xliff命名空间
    if 'xliff:' in content and 'xmlns:xliff' not in content:
        content = content.replace('<resources>',
                                  '<resources xmlns:xliff="urn:oasis:names:tc:xliff:document:1.2">', 1)

    # 写入文件（保留原始格式）
    os.makedirs(os.path.dirname(xml_path), exist_ok=True)
    with open(xml_path, 'w', encoding='utf-8') as f:
        f.write(content)


"""
导出 strings.xml 中的内容到 Excel 表格：python3 processor.py --export res_dir translations.xlsx
将 Excel 表格中的内容导入到 strings.xml：python3 processor.py --import res_dir translations.xlsx

positional arguments:
  res_dir     资源目录路径
  excel_file  Excel文件路径

options:
  -h, --help  show this help message and exit
  --export    导出到Excel
  --import    从Excel导入
  
"""
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Android字符串精确处理器')
    parser.add_argument('--export', action='store_true', help='导出到Excel')
    parser.add_argument('--import', action='store_true', dest='import_', help='从Excel导入')
    parser.add_argument('res_dir', help='资源目录路径')
    parser.add_argument('excel_file', help='Excel文件路径')

    args = parser.parse_args()

    if args.export:
        export_to_excel(args.res_dir, args.excel_file)
    elif args.import_:
        import_from_excel(args.res_dir, args.excel_file)
    else:
        print("请使用--export或--import参数")
