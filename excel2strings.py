
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
            key_text = keys[index]
            value_text = values_en[index]
            # 如果KEY内容为空(""),将被跳过
            if key_text != "":
                text_line = "\"%s\" = \"%s\"" % (key_text, value_text)
                file_object.write(text_line + ";\n")
                
    # 写入文件 SC
    filename_sc = 'output_' + sheet_name + '_sc.strings'
    with open(filename_sc, 'w') as file_object:
        for index in range(len(keys)):
            key_text = keys[index]
            value_text = values_sc[index]
            # 如果KEY内容为空(""),将被跳过
            if key_text != "":
                text_line = "\"%s\" = \"%s\"" % (key_text, value_text)
                file_object.write(text_line + ";\n")

    # 写入文件 TC
    filename_tc = 'output_' + sheet_name + '_tc.strings'
    with open(filename_tc, 'w') as file_object:
        for index in range(len(keys)):
            key_text = keys[index]
            value_text = values_tc[index]
            # 如果KEY内容为空(""),将被跳过
            if key_text != "":
                text_line = "\"%s\" = \"%s\"" % (key_text, value_text)
                file_object.write(text_line + ";\n")
