#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
import time
import xml
import xml.dom.minidom
from distutils.log import Log
from optparse import OptionParser
from xml.etree import ElementTree

import xlwt
from bs4 import BeautifulSoup


def read_xml(path):
    """通过ElementTree获取

    大部分字符串内容都能读取出来，但是如果<string></string>标签内嵌套了子标签，
    那么子标签内的内容读取不出来，并且只能读取到第一个子标签前的内容

    :param path:
    :return:
    """
    if path is None or len(path) == 0:
        Log().error('file path is None')
        return

    file = open(path, encoding='utf-8')
    string_list = file.read()
    root = ElementTree.fromstringlist(string_list)
    item_list = root.findall('string')
    keys = []
    values = []
    for item in item_list:
        key = item.attrib['name']
        value = item.text
        keys.append(key)
        values.append(value)
    file.close()
    return keys, values


def read_xml2(path):
    """通过DOM获取

    如果<string></string>标签内紧跟着子标签, 比如<string><font></font></string>
    ，则会报错，提示item.firstChild.data这句话没有data属性;
    如果<string></string>标签内先是内容然后跟着子标签，则会读取第一个子标签前的内容

    :param path:
    :return:
    """
    if path is None or len(path) == 0:
        Log().error('file path is None')
        return

    dom = xml.dom.minidom.parse(path)
    root = dom.documentElement
    item_list = root.getElementsByTagName('string')
    keys = []
    values = []
    for index in range(len(item_list)):
        item = item_list[index]
        key = item.getAttribute("name")
        value = item.firstChild.data
        keys.append(key)
        values.append(value)
    return keys, values


def read_xml3(path):
    """通过BeautifulSoup获取

    如果<string></string>标签内含子标签，则子标签内的内容也可读取出来，
    注意：如果使用的是BeautifulSoup(file, 'xml')，则文件里的数据会出现读取不全的情况，可能只读取了几个<string></string>内的数据

    :param path:
    :return:
    """
    if path is None or len(path) == 0:
        Log().error('file path is None')
        return None, None

    file = open(path)
    soup = BeautifulSoup(file, 'lxml')
    strings = soup.findAll('string')
    keys = []
    values = []
    for string in strings:
        key = string.get('name')
        # value = string.string 如果包含多个子标签， 结果返回None
        value = del_content_blank(string.get_text().strip())
        keys.append(key)
        values.append(value)
    file.close()
    return keys, values


def del_content_blank(s):
    clean_str = re.sub(r'\n| {8}', ' ', str(s))
    return clean_str.replace('  ', ' ')


def get_country_code(dir_name):
    code = 'en'
    dir_split = dir_name.split('values-')
    if len(dir_split) > 1:
        code = dir_split[1]
    return code


def get_dest_dir(target_dir, option):
    if option == 1:
        dir = 'single_file_one_sheet'
    elif option == 2:
        dir = 'single_file_multiple_sheets'
    elif option == 3:
        dir = 'multiple_files'
    else:
        dir = 'multiple_files_no_translate'
    dest_dir = target_dir + "/xml2xls/" + dir + '/' + time.strftime("%Y%m%d_%H%M%S")
    if not os.path.exists(dest_dir):
        os.makedirs(dest_dir)
    return dest_dir


def convert_to_multiple_files(file_dir, target_dir):
    dest_dir = get_dest_dir(target_dir, 3)
    for _, dir_names, _ in os.walk(file_dir):
        values_dirs = [di for di in dir_names if di.startswith("values")]
        for dir_name in values_dirs:
            country_code = get_country_code(dir_name)
            xml_file = 'strings.xml'
            xml_file_path = file_dir + '/' + dir_name + '/' + xml_file
            if not os.path.exists(xml_file_path):
                continue
            file_name = xml_file.replace(".xml", "-" + country_code)
            sheet_name = file_name
            dest_file_path = dest_dir + "/" + file_name + ".xls"
            if not os.path.exists(dest_file_path):
                workbook = xlwt.Workbook(encoding='utf-8')
                ws = workbook.add_sheet(sheet_name)
                ws.write(0, 0, 'name')
                ws.write(0, 1, country_code)
                path = file_dir + '/' + dir_name + '/' + xml_file
                (keys, values) = read_xml3(path)

                print("Start Converting %s " % country_code)
                print('Total: %s' % len(keys))

                for x in range(len(keys)):
                    key = keys[x]
                    value = values[x]
                    ws.write(x + 1, 0, key)
                    ws.write(x + 1, 1, value)
                workbook.save(dest_file_path)
                print("Convert %s successfully! you can see xls file in %s" % (path, dest_dir))


def convert_to_single_file_with_one_sheet(file_dir, target_dir):
    dest_dir = get_dest_dir(target_dir, 1)
    sheet_name = 'strings'
    dest_file_path = dest_dir + "/" + "single_file_one_sheet.xls"
    if not os.path.exists(dest_file_path):
        workbook = xlwt.Workbook(encoding='utf-8')
        ws = workbook.add_sheet(sheet_name)
        ws.write(0, 0, 'name')
        for _, dir_names, _ in os.walk(file_dir):
            values_dirs = [di for di in dir_names if di.startswith("values")]
            index = 0
            en_keys = []
            values_dirs.sort()
            for dir_name in values_dirs:
                xml_file = 'strings.xml'
                xml_file_path = file_dir + '/' + dir_name + '/' + xml_file
                if not os.path.exists(xml_file_path):
                    continue
                country_code = get_country_code(dir_name)
                ws.write(0, index + 1, country_code)
                (keys, values) = read_xml3(xml_file_path)

                print("Start Converting %s " % country_code)
                print('Total: %s' % len(keys))

                if country_code == 'en':
                    en_keys = keys
                    for x in range(len(keys)):
                        key = keys[x]
                        value = values[x]
                        ws.write(x + 1, 0, key)
                        ws.write(x + 1, 1, value)
                else:
                    for x in range(len(keys)):
                        key = keys[x]
                        for x2 in range(len(en_keys)):
                            key2 = en_keys[x2]
                            if key == key2:
                                value = values[x]
                                ws.write(x2 + 1, index + 1, value)
                print("Convert %s successfully! you can see xls file in %s" % (xml_file_path, dest_dir))
                index += 1
        workbook.save(dest_file_path)


def convert_to_single_file_with_multiple_sheets(file_dir, target_dir):
    dest_dir = get_dest_dir(target_dir, 2)
    dest_file_path = dest_dir + "/" + "single_file_multi_sheets.xls"
    workbook = xlwt.Workbook(encoding='utf-8')
    for _, dirnames, _ in os.walk(file_dir):
        values_dirs = [di for di in dirnames if di.startswith("values")]
        for dirname in values_dirs:
            xml_file = 'strings.xml'
            xml_file_path = file_dir + '/' + dirname + '/' + xml_file
            if not os.path.exists(xml_file_path):
                continue
            country_code = get_country_code(dirname)
            sheet_name = xml_file.replace(".xml", "-" + country_code)
            if not os.path.exists(dest_file_path):
                ws = workbook.add_sheet(sheet_name)
                ws.write(0, 0, 'name')
                ws.write(0, 1, country_code)
                path = file_dir + '/' + dirname + '/' + xml_file
                (keys, values) = read_xml3(path)

                print('Start Converting %s' % country_code)
                print('Total: %s' % len(keys))

                for x in range(len(keys)):
                    key = keys[x]
                    value = values[x]
                    ws.write(x + 1, 0, key)
                    ws.write(x + 1, 1, value)
                print("Convert %s successfully! you can see xls file in %s" % (path, dest_dir))
    workbook.save(dest_file_path)


def convert_to_multiple_files_no_translate(file_dir, target_dir):
    dest_dir = get_dest_dir(target_dir, 4)
    for _, dir_names, _ in os.walk(file_dir):
        values_dirs = [di for di in dir_names if di.startswith("values")]
        values_dirs.sort()
        en_keys = []
        en_values = []
        for dir_name in values_dirs:
            country_code = get_country_code(dir_name)
            xml_file = 'strings.xml'
            xml_file_path = file_dir + '/' + dir_name + '/' + xml_file
            if not os.path.exists(xml_file_path):
                continue
            path = file_dir + '/' + dir_name + '/' + xml_file
            (keys, values) = read_xml3(path)
            if country_code == 'en':
                en_keys = keys
                en_values = values
            else:

                print("Start converting %s" % country_code)
                print('Translated Count: %s' % len(keys))

                file_name = xml_file.replace(".xml", "_no_translate_to_" + country_code)
                sheet_name = file_name
                dest_file_path = dest_dir + "/" + file_name + ".xls"
                if not os.path.exists(dest_file_path):
                    workbook = xlwt.Workbook(encoding='utf-8')
                    ws = workbook.add_sheet(sheet_name)
                    ws.write(0, 0, 'name')
                    ws.write(0, 1, 'en')
                    index = 0
                    for x in range(len(en_keys)):
                        key = en_keys[x]
                        if key not in keys:
                            value = en_values[x]
                            ws.write(index + 1, 0, key)
                            ws.write(index + 1, 1, value)
                            index += 1
                    workbook.save(dest_file_path)
                    print("Untranslated Count: %s" % index)
                    print("Convert %s successfully! you can see xls file in %s" % (path, dest_dir))


def add_parser():
    # usage = "xml2xls.py -f fileDir -t targetDir -e excelStorageForm"
    # parser = OptionParser(usage=usage)

    parser = OptionParser()

    parser.add_option("-f", "--fileDir",
                      help="The parent directory of values(values, values-es, values-pt, ...)directory",
                      metavar="fileDir")

    parser.add_option("-t", "--targetDir",
                      help="The directory where the xls files will be saved.",
                      metavar="targetDir")

    parser.add_option("-e", "--excelStorageForm",
                      type="int",
                      default="1",
                      help="The excel(.xls) file storage forms including single file with one sheet(-e 1), single file "
                           "with multiple sheet(-e 2), multiple files(-e 3), multiple files with untranslated(-e 4), "
                           "or above all(-e 5). "
                           "Default is single file with one sheet(-e 1).",
                      metavar="excelStorageForm")

    (options, args) = parser.parse_args()
    # print("options: %s, args: %s" % (options, args))

    return options


def start_convert(options):
    file_dir = options.fileDir
    target_dir = options.targetDir

    if file_dir is None:
        Log().error("strings.xml files directory can not be empty! try -h for help.")
        return

    if not os.path.exists(file_dir):
        Log().error("%s does not exist." % file_dir)
        return

    if target_dir is None:
        target_dir = os.getcwd()

    print("------------------------------Start converting------------------------------")

    if options.excelStorageForm == 1:
        convert_to_single_file_with_one_sheet(file_dir, target_dir)
    elif options.excelStorageForm == 2:
        convert_to_single_file_with_multiple_sheets(file_dir, target_dir)
    elif options.excelStorageForm == 3:
        convert_to_multiple_files(file_dir, target_dir)
    elif options.excelStorageForm == 4:
        convert_to_multiple_files_no_translate(file_dir, target_dir)
    elif options.excelStorageForm == 5:
        convert_to_single_file_with_one_sheet(file_dir, target_dir)
        convert_to_single_file_with_multiple_sheets(file_dir, target_dir)
        convert_to_multiple_files(file_dir, target_dir)
        convert_to_multiple_files_no_translate(file_dir, target_dir)
    else:
        Log().error('Invalid value %s , -e only for values 1, 2, 3, 4, 5' % options.excelStorageForm)


def main():
    options = add_parser()
    start_convert(options)
    # convert_to_single_file_with_multiple_sheets('/home/shewenbiao/Android/Workspace/CompanyProject/CleanMaster/Cleaner/app/src/main/res', os.getcwd())
    # convert_to_multiple_files('/home/shewenbiao/Android/Workspace/CompanyProject/CleanMaster/Cleaner/app/src/main/res', os.getcwd())
    # convert_to_single_file_with_one_sheet('/home/shewenbiao/Android/Workspace/CompanyProject/CleanMaster/Cleaner/app/src/main/res', os.getcwd())
    # convert_to_multiple_files_no_translate('/home/shewenbiao/Android/Workspace/CompanyProject/CleanMaster/Cleaner/app/src/main/res', os.getcwd())


main()
