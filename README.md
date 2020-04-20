# excel2strings

A tiny tool to convert excel file to iOS .strings files.

## Dependencies

- XLRD  <https://github.com/python-excel/xlrd>
  
    > Run `pip install xlrd`

## Usage

1. Put excel2strings.py & excel file into the same file folder.
2. Edit `excel2strings.py`, and specify your initialization parameters.
3. Run `Python3 excel2strings.py`, then it will output some `.strings` files.

## Tips

If you want to combine all `xxx_en.strings` file, just excute shell `cat 1_en.strings 2_en.string > en.strings`.
