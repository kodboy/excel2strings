# excel2strings
一个Excel转.strings文件的小工具

* 遍历多个sheet
* 去重,重复的key,后值替换前值

## 依赖库

- openpyxl
  
    > 运行 `pip install openpyxl`

## 使用方法

1. 把 `excel2strings.py` 和 excel 文件放到同一个文件目录
2. 编辑 `excel2strings.py`, 指定初始化参数
3. 运行 `Python3 excel2strings.py`, 会自动导出一些.strings文件和log文件
