# strings2xls
Python command line tool for conversion between android strings.xml files and excel files.

## 最新脚本(更新日期：2025/03/13)：processor.py

### Change log
1. 该脚本使用起来更方便。导出和导入使用这一个脚本即可。
2. 解决了老版本中一些因字符串内容包含特殊标签（比如 Html 标签）导致导出到表格中特殊标签缺失，或者内容缺失的问题。
3. 该脚本导出到表格包含两部分：一部分是将全部语言下的字符串导出到一个 sheet（All Translations）, 另外一部分是所有未翻译的字符串会放到一个 sheet（Untranslated）中。

### 命令
1. 导出 strings.xml 中的内容到 Excel 表格
```
python3 processor.py --export res_dir excel_file
```
2. 将 Excel 表格中的内容导入到 strings.xml
```
python3 processor.py --import res_dir excel_file --mode [full|partial]
```
### 使用示例
1. 导出所有翻译（含未翻译项）：
   ```
   python3 processor.py --export app/src/main/res translations.xlsx
   ```

2. 导入完整翻译表（All Translations）数据：
   ```
   python3 processor.py --import app/src/main/res translations.xlsx --mode full
   ```

3. 仅导入未翻译表（Untranslated）数据：
   ```
   python3 processor.py --import app/src/main/res translations.xlsx --mode partial
   ```
   

### 参数说明
```
positional arguments:
  res_dir               资源目录路径（包含 values/values-xx 的文件夹）
  excel_file            Excel文件路径（输入/输出）

optional arguments:
  -h, --help            show this help message and exit
  --export              导出到Excel（生成完整翻译表和未翻译项表）
  --import              从Excel导入（需配合 --mode 选择数据源, 支持两种模式）
  --mode {full,partial}
                        
                        导入模式选择：                      
                          full - 从主表(All Translations)导入（默认）
                          partial - 仅从未翻译表(Untranslated)导入
```                          


### 安装 openpyxl
```
pip3 install openpyxl
```

## 特性

- [x] 支持将 **android strings** xml 文件转换成 **excel** 文件
- [x] 支持将 **excel** 文件转换成 **android strings** xml 文件

## 所需环境

### 以下以Ubuntu 16.0.4为例

### 1.安装python3

python 版本必须是 3.x

安装python3.6为例:
```
$ sudo add-apt-repository ppa:jonathonf/python-3.6
$ sudo apt-get update
$ sudo apt-get install python3.6
$ python3 --version
Python 3.6.3
```
如果已经安装了python3.x其他版本，以下几个安装步骤也是一样

### 2.检查 pip3(python3 包管理器)

```
$ pip3 --version
pip 19.3.1 from /usr/local/lib/python3.6/site-packages/pip (python 3.6)
```

如果没有安装 pip3

```
$ sudo apt-get update
$ sudo apt-get install python3-pip
```

### 3.安装 xlwt

```
$ sudo pip3 install xlwt
```

### 4.安装 xlrd

```
$ sudo pip3 install xlrd
```

### 5.安装 beatifulsoup4

```
$ sudo pip3 install beautifulsoup4
```
### 6.安装 lxml

```
$ sudo pip3 install lxml
```

## 使用说明
### 1.将 **android strings** xml 文件转换成 **excel** 文件

命令
```
$ python3 xml2xls.py -f fileDir -t targetDir -e excelStorageForm
```
fireDir: 项目的res目录路径或者其他目录路径(该目录里需包含values目录或者不同语言的values目录（比如values-es, values-pt等等)，在各个values目录下放各自的strings.xml文件，需要注意的是该文件名只能是strings.xml)。最简单的是直接指定项目的res目录路径。

targetDir: 转换后的xls表格保存的目录路径。不指定的话，默认保存在当前目录下。

excelStorageForm: 可选的值有1, 2, 3, 4, 5。不指定的话，采用默认值1。各个值代表的含义如下

1: 将所有语言下的字符串输出到同一个表格的同一个sheet里

2: 将所有语言下的字符串输出到同一个表格的多个sheet里，一个sheet对应一种语言

3: 将所有语言下的字符串输出到多个表格里，一个表格对应一种语言

4: 将默认语言下的字符串在其他语言下没有翻译的字符串输出到多个表格里，一个表格对应一种语言

5: 包含以上4种


### 2.将 **excel** 文件转换成 **android strings** xml 文件

命令
```
$ python3 xls2xml.py -f fileDir -t targetDir
```
fileDir: 要转换的xls表格所在的目录路径

targetDir: 转换后的xml文件保存的目录路径。不指定的话，默认保存在当前目录下。

