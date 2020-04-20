import xlrd

########################################### 初始化定义 - 开始
# 指定Excel文件名, (别忘加后缀)
excel_file_name = "nls.xlsx"
# 指定sheets名字s
sheet_names = ["Reset secondary password", "Reset password", "Error handling", "Device optional page"]
key_prefixs = ["reset_password_with_", "reset_secondary_password_with_", "reset_password_error_handing_with_", "device_optional_with_"]

#### 指定关键行号
# 标题行的行号(没用使用)
rowIndexForTitle = 0
# 首行 Key-Value 的行号 (标题行)
rowIndexForStartKeyValue = 1

#### 指定关键列号
# Key的列号
columnIndexForAllKeys = 1
# 英文列的列号
columnIndexForValues_en = 2
# 简体中文列的列号
columnIndexForValues_sc = 4
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
    values_en = sheet.col_values(columnIndexForValues_en, rowIndexForStartKeyValue)
    values_sc = sheet.col_values(columnIndexForValues_sc, rowIndexForStartKeyValue)
    values_tc = sheet.col_values(columnIndexForValues_tc, rowIndexForStartKeyValue)

    # 写入文件 EN
    filename_en = 'output_' + sheet_name + '_en.strings'
    with open(filename_en, 'w') as file_object:
        for index in range(len(keys)):
            key_prefix = key_prefixs[index]
            key_text = keys[index]
            value_text = values_en[index]
            # 如果KEY内容为空(""),将被跳过
            if key_text != "":
                text_line = "\"%s%s\" = \"%s\"" % (key_prefix,key_text, value_text)
                file_object.write(text_line + ";\n")
                
    # 写入文件 SC
    filename_sc = 'output_' + sheet_name + '_sc.strings'
    with open(filename_sc, 'w') as file_object:
        for index in range(len(keys)):
            key_prefix = key_prefixs[index]
            key_text = keys[index]
            value_text = values_sc[index]
            # 如果KEY内容为空(""),将被跳过
            if key_text != "":
                text_line = "\"%s%s\" = \"%s\"" % (key_prefix,key_text, value_text)
                file_object.write(text_line + ";\n")

    # 写入文件 TC
    filename_tc = 'output_' + sheet_name + '_tc.strings'
    with open(filename_tc, 'w') as file_object:
        for index in range(len(keys)):
            key_prefix = key_prefixs[index]
            key_text = keys[index]
            value_text = values_tc[index]
            # 如果KEY内容为空(""),将被跳过
            if key_text != "":
                text_line = "\"%s%s\" = \"%s\"" % (key_prefix,key_text, value_text)
                file_object.write(text_line + ";\n")

