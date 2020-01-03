# strings2xls
Python command line tool for conversion between android strings.xml files and excel files.Reference https://github.com/CatchZeng/Localizable.strings2Excel 基于该链接里的基础上做了改善，以及扩展。


## 特性

- [x] 支持将 **android** xml 文件转换成 **excel** 文件
- [x] 支持将 **excel** 文件转换成 **android** xml 文件

## 所需环境

### 以下以Ubuntu 16.0.4为例

### 1.安装python3

python 版本必须是 3.x

```
$ sudo add-apt-repository ppa:jonathonf/python-3.6
$ sudo apt-get update
$ sudo apt-get install python3.6
$ python3 --version
Python 3.6.3
```

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
sudo pip3 install xlwt
```

### 4.安装 xlrd

```
sudo pip3 install xlrd
```

### 5.安装 beatifulsoup4

```
sudo pip3 install beautifulsoup4
```
### 6.安装 lxml

```
sudo pip3 install lxml
```

## 使用说明
### 1.将 **android** xml 文件转换成 **excel** 文件



### 2.将 **excel** 文件转换成 **android** xml 文件




