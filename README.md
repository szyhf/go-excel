# Intro | 简介

Expect to create a ORM-Like library to read or write relate-db-like excel easily.

---

在复杂的系统中（例如游戏）

有时候为了便于非专业人员设置一些配置

会使用Excel作为一种轻量级的关系数据库或者配置文件

毕竟对于很多非开发人员来说

配个Excel要比写json或者yaml什么简单得多

这种场景下

读取特定格式（符合关系数据库特点的表格）的数据会比各种花式写入Excel的功能更重要

毕竟从编辑上来说微软提供的Excel本身功能就非常强大了

而现在我找到的Excel库的功能都过于强大了

用起来有点浪费

于是写了这个简化库

这个库的工作参考了[tealeg/xlsx](github.com/tealeg/xlsx)的部分实现和读取逻辑。

感谢[tealeg](github.com/tealeg)

## RoadMap | 开发计划

+ Read xlsx file and got the expect xml. √
+ Prepare the shared string xml. √
+ Get the correct sheetX.xml. √
+ Read a row of a sheet. √
+ Read a cell of a row, fix the empty cell. √
+ Fill string cell with value of shared string xml. √
+ Read a row to a struct by column index.
+ Can set the column name row, default is the first row.
+ Read a row to a struct by column name.
+ To be continued...

