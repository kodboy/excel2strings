# excel2strings
一个Excel转.strings文件的小工具

## 依赖库

- XLRD  <https://github.com/python-excel/xlrd>
  
    > Run `pip install xlrd`

## 使用方法

1. 把 `excel2strings.py` 和 excel 文件放到同一个文件目录
2. 编辑 `excel2strings.py`, 指定初始化参数
3. 运行 `Python3 excel2strings.py`, 会自动导出一些.strings文件

## 小提示

如果想合并所有`xxx_en.strings`文件,直接执行shell命令
1. 可能引起文件乱序
```shell
cat *_en.strings > en.strings
```
或者使用 2. 给定顺序
```
cat xxx_en.strings yyy_en.string > en.strings
```
