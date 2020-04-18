import xlrd

########################################### 初始化定义 - 开始
# 指定Excel文件名, (别忘加后缀)
excel_file_name = "income.xlsx"
# 指定sheets名字s
sheet_names = ["2018", "2017", "2016"]

#### 指定关键行号
# 标题行的行号(没用使用)
rowIndexForTitle = 0
# 首行 Key-Value 的行号 !important
rowIndexForStartKeyValue = 1

#### 指定关键列号
# Key的列号
columnIndexForAllKeys = 0
# 英文列的列号
columnIndexForValues_en = 1
# 简体中文列的列号
columnIndexForValues_sc = 2
# 繁体中文列的列号
columnIndexForValues_tc = 3
############################################# 初始化定义 - 结束

book = xlrd.open_workbook(excel_file_name)
# 遍历sheets
for sheet_name in sheet_names:
    # 获取sheet
    sheet = book.sheet_by_name(sheet_name)
    # all keys string(s)
    keys = sheet.col_values(columnIndexForAllKeys, rowIndexForStartKeyValue)
    # all english value string(s)
    values_EN = sheet.col_values(columnIndexForValues_en, rowIndexForStartKeyValue)
    values_ZH_CN = sheet.col_values(columnIndexForValues_en, rowIndexForStartKeyValue)
    values_TS_CN = sheet.col_values(columnIndexForValues_sc, rowIndexForStartKeyValue)

    # 写入文件 EN
    filename_en = 'output_' + sheet_name + '_en.strings'
    with open(filename_en, 'w') as file_object:
        for index in range(len(keys)):
            text_line = "\"%s\" = \"%s\"" % (keys[index], values_EN[index])
            file_object.write(text_line + "\n")

    # 写入文件 SC
    filename_sc = 'output_' + sheet_name + '_sc.strings'
    with open(filename_sc, 'w') as file_object:
        for index in range(len(keys)):
            text_line = "\"%s\" = \"%s\"" % (keys[index], values_ZH_CN[index])
            file_object.write(text_line + "\n")

    # 写入文件 TC
    filename_tc = 'output_' + sheet_name + '_tc.strings'
    with open(filename_tc, 'w') as file_object:
        for index in range(len(keys)):
            text_line = "\"%s\" = \"%s\"" % (keys[index], values_TS_CN[index])
            file_object.write(text_line + "\n")


