# strings2xls
Python command line tool for conversion between android strings.xml files and excel files.Reference https://github.com/CatchZeng/Localizable.strings2Excel 基于该链接里的基础上做了改善，以及扩展。


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

2: 将所有语言下的字符传输出到同一个表格的多个sheet里，一个sheet对应一种语言

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

