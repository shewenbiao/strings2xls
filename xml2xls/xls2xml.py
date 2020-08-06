#!/usr/bin/env python3
# -*- coding: utf-8 -*-


from distutils.log import Log
from optparse import OptionParser

import xlrd
import os

import time


def open_excel(path):
    try:
        data = xlrd.open_workbook(path, encoding_override="utf-8")
        return data
    except Exception as ex:
        return ex


def read_from_excel(file_path):
    data = open_excel(file_path)
    table = data.sheets()[0]
    keys = table.col_values(0)
    del keys[0]
    # print(keys)
    first_row = table.row_values(0)
    lan_values = {}
    for index in range(len(first_row)):
        if index <= 0:
            continue
        language_name = first_row[index]
        # print(language_name)
        values = table.col_values(index)
        del values[0]
        # print(values)
        lan_values[language_name] = values
    return keys, lan_values


def write_to_xml(keys, values, file_path, language_name):
    fo = open(file_path, "wb")
    string_encoding = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n<resources>\n"
    fo.write(bytes(string_encoding, encoding="utf8"))
    for x in range(len(keys)):
        if values[x] is None or values[x] == '':
            Log().error("Language: " + language_name + " Key:" + keys[x] + " value is None. Index:" + str(x + 1))
            continue
        key = keys[x].strip()
        # value = re.sub(r'(%\d\$)(@)', r'\1s', str(values[x]))
        value = str(values[x])
        content = "    <string name=\"" + key + "\">" + value + "</string>\n"
        fo.write(bytes(content, encoding="utf8"))
    fo.write(bytes("</resources>", encoding="utf8"))
    fo.close()


def add_parser():
    parser = OptionParser()

    parser.add_option("-f", "--fileDir",
                      help="Xls files directory.",
                      metavar="fileDir")

    parser.add_option("-t", "--targetDir",
                      help="The directory where the xml files will be saved.",
                      metavar="targetDir")

    (options, args) = parser.parse_args()
    # print("options: %s, args: %s" % (options, args))

    return options


def convert_to_xml(file_dir, target_dir):
    dest_dir = target_dir + "/xls2xml/" + time.strftime("%Y%m%d_%H%M%S")
    for _, _, file_names in os.walk(file_dir):
        xls_file_names = [fi for fi in file_names if fi.endswith(".xls") or fi.endswith(".xlsx")]
        for file in xls_file_names:
            data = xlrd.open_workbook(file_dir + "/" + file, 'utf-8')
            sheet = data.sheets()
            for table in sheet:
                first_row = table.row_values(0)
                keys = table.col_values(0)
                del keys[0]

                for index in range(len(first_row)):
                    if index <= 0:
                        continue
                    language_name = first_row[index]
                    values = table.col_values(index)
                    del values[0]

                    if language_name == "zh-Hans":
                        language_name = "zh-rCN"

                    path = dest_dir + "/values-" + language_name + "/"
                    if language_name == 'en':
                        path = dest_dir + "/values/"
                    if not os.path.exists(path):
                        os.makedirs(path)
                    filename = 'strings.xml'
                    write_to_xml(keys, values, path + filename, language_name)
    print("Convert %s successfully! you can xml files in %s" % (
        file_dir, dest_dir))


def start_convert(options):
    file_dir = options.fileDir
    target_dir = options.targetDir

    print("Start converting")

    if file_dir is None:
        Log().error("xls files directory can not be empty! try -h for help.")
        return

    if not os.path.exists(file_dir):
        Log().error("%s does not exist." % file_dir)
        return

    if target_dir is None:
        target_dir = os.getcwd()

    if not os.path.exists(target_dir):
        os.makedirs(target_dir)

    convert_to_xml(file_dir, target_dir)


def main():
    options = add_parser()
    start_convert(options)

    # convert_to_xml("/Users/shewenbiao/Desktop/xls2xml", os.getcwd())


main()
