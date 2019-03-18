# Pyqt5ExcelToMysql
使用PyQt5将Excel数据导入mysql

# 使用PyQt5将Excel数据导入mysql

[TOC]

## 一、环境配置

系统：win10,64位，ltsc2019

工具：pycharm 2018.03.04

<!--more-->

语言：python 3.5.4

## 二、使用的库函数



PyQt5  `pip install pyqt5`

pymysql `pip install pymysql`

xlrd	`pip install xlrd`

## 三、QT Designer设计界面

![](https://img.rayu.me/2019/03/359316014.png)

![](https://img.rayu.me/2019/03/3198127376.png)

具体配置见UI文件。

设计完成后，使用pyUIC将UI文件转换为python代码。

可以使用pycharm调用外部工具：
pyuic.exe的目录在python的安装目录下，working directory:内容填写`$FileDir$`

![](https://img.rayu.me/2019/03/2259217360.png)

然后使用该工具进行转换

![](https://img.rayu.me/2019/03/1732376612.png)

转换完成后，工程目录下会有一个project.py，名称同UI文件。

## 四、代码编写

> 参考链接：[PyQT5练习：制作Excel文件导入MySQL窗口](https://zhuanlan.zhihu.com/p/54867294)

参考链接中的评论说用xlrd速度很快，于是尝试之。
导入库函数，和定义一个类及一些变量
![](https://img.rayu.me/2019/03/1294301532.png)

初始化界面按钮，默认为false，绑定按钮到对应的函数
![初始化按钮](https://img.rayu.me/2019/03/856615813.png)

连接数据库函数
![连接数据库](https://img.rayu.me/2019/03/4048418983.png)

断开连接函数
![断开](https://img.rayu.me/2019/03/2914372686.png)

打开文件函数
![打开](https://img.rayu.me/2019/03/137168773.png)

设置选中sheet名字后，导入按钮才可用
![](https://img.rayu.me/2019/03/86425240.png)

写入数据库，获取表名，将sheetname作为数据库中的表名
![](https://img.rayu.me/2019/03/209013864.png)

整理建表SQL语句，默认为Excel中的第一行为字段，如果表不存在的情况。

![](https://img.rayu.me/2019/03/2677207670.png)

整理插入SQL语句。
![](https://img.rayu.me/2019/03/1276142188.png)

写入数据库并计算插入时间
![](https://img.rayu.me/2019/03/1774937210.png)

处理关闭按钮事件，点击关闭时断开数据库连接
![](https://img.rayu.me/2019/03/761226418.png)

主函数
![](https://img.rayu.me/2019/03/2230780151.png)

## 五、运行结果
界面比较LOW，请忽略
![](https://img.rayu.me/2019/03/2663345438.png)

插入10000多条数据的时间，2秒多。
![](https://img.rayu.me/2019/03/2234676904.png)

## 六、TODO list
[] 1、自定义表名和表内字段名称；
