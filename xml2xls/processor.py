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
    """高精度解析器，保留完整原始内容"""
    order = []
    strings = {}
    if not os.path.exists(xml_path):
        return order, strings

    with open(xml_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # 匹配完整字符串内容（包含CDATA和普通标签）
    pattern = re.compile(
        r'<string\s+name="([^"]+)">(.*?)</string>',
        re.DOTALL
    )

    for match in re.finditer(pattern, content):
        name = match.group(1)
        raw_value = match.group(2).strip()

        # 保留完整的原始内容
        strings[name] = raw_value
        order.append(name)

    return order, strings


def write_strings_xml(xml_path, data):
    """写入XML时保留原始格式"""
    os.makedirs(os.path.dirname(xml_path), exist_ok=True)

    with open(xml_path, 'w', encoding='utf-8') as f:
        f.write('<?xml version="1.0" encoding="utf-8"?>\n<resources>\n')

        for name, value in data.items():
            line = f'    <string name="{name}">{value}</string>'
            f.write(line + '\n')

        f.write('</resources>')


"""
导出 strings.xml 中的内容到 Excel 表格：python3 processor.py --import res_dir translations.xlsx
将 Excel 表格中的内容导入到 strings.xml：python3 processor.py --export res_dir translations.xlsx

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
