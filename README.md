# excel2strings

A tiny tool to convert excel file to iOS .strings files.
  - `excel2strings.py`,没去重,提示重复信息
  - `convert_no_repeat.py`,转换并去重

## Dependencies

- XLRD  <https://github.com/python-excel/xlrd>
> Run `pip install xlrd`

## Usage

1. Put py file & excel file into the same file folder.
2. Edit py file, and specify your initialization parameters.
3. Run `Python3 excel2strings.py` or `Python3 convert_no_repeat.py`, then it will output some `.strings` files.
