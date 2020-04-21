import xlrd

# log重复KEY信息,保留不去重的strings和去重后的strings

########################################### 初始化定义 - 开始
# 指定Excel文件名, (别忘加后缀)
excel_file_name = "nls.xlsx"
# 指定sheets名字s
sheet_names = ["Reset secondary password", "Reset password", "Error handling", "Device optional page"]
sheet_key_prefixs = ["", "", "", ""]

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

map_record = {} # Check en only

# 为了保证key的导出顺序,用list来记录key的出现顺序,如果有key重复出现,将会移除前一个key. (另外,经测试,从map遍历key-value导出也是同一个结果,以防万一乱序)
# 去重后的有序 keys
keys_order_list = [] 
# 去重后的key-value,全部sheet遍历完成后再写入
map_en = {}
map_sc = {}
map_tc = {}

# excel对象
book = xlrd.open_workbook(excel_file_name)

# 遍历sheets
for sheet_index in range(len(sheet_names)):
    sheet_name = sheet_names[sheet_index]
    key_prefix = sheet_key_prefixs[sheet_index]
    
    # 获取sheet对象
    sheet = book.sheet_by_name(sheet_name)
    # all keys string(s)
    keys = sheet.col_values(columnIndexForAllKeys, rowIndexForStartKeyValue)
    # all values string(s)
    values_en = sheet.col_values(columnIndexForValues_en, rowIndexForStartKeyValue)
    values_sc = sheet.col_values(columnIndexForValues_sc, rowIndexForStartKeyValue)
    values_tc = sheet.col_values(columnIndexForValues_tc, rowIndexForStartKeyValue)
    
    # 写入文件 EN
    filename_en = sheet_name + '_en.strings'
    with open("en.strings", 'a+') as file_en:
        with open(filename_en, 'w') as file_object:
            # file_object.write("/* %s */\n" % (sheet_name))
            print("filename_en: %s" % filename_en)
            for index in range(len(keys)):
                key_text = keys[index]
                if key_text != "": # 如果KEY内容为空(""),将被跳过
                    value_text = values_en[index].replace("\"", "\\\"").replace("\n", "\\n").replace("\\n\"", "\"").lstrip()
                    map_new_record = "sheet:%s:%s" % (sheet_name, (index + rowIndexForStartKeyValue + 1)) # 定位到原文件中sheet+行号
                    map_en[key_text] = value_text
                    if keys_order_list.count(key_text) != 0:
                        keys_order_list.remove(key_text)
                    keys_order_list.append(key_text)
                    # 处理提示
                    if map_record.get(key_text) is None:  # 无重复
                        map_record[key_text] = map_new_record # 加入 map_record
                    else:
                        map_old_record = map_record.get(key_text)
                        map_record[key_text] = "%s ; %s" % (map_old_record,map_new_record)
                        with open("log.strings", 'a+') as log:
                            log.write("------------------------------\n")
                            log.write("Duplicate Key:%s\n" % key_text)
                            log.write("%s\n\n" % map_record[key_text])
                    en_new_value = "\"%s%s\" = \"%s\"" % (key_prefix,key_text, value_text) # 加入 map_key_value_en
                    file_object.write(en_new_value + ";\n")
                    file_en.write(en_new_value + ";\n")
                        
    # 写入文件 SC
    filename_sc = sheet_name + '_sc.strings'
    with open("sc.strings", 'a+') as file_sc:
        with open(filename_sc, 'w') as file_object:
            # file_object.write("/* %s */\n" % (sheet_name))
            print("filename_sc: %s" % filename_sc)
            for index in range(len(keys)):
                key_text = keys[index]
                # 如果KEY内容为空(""),将被跳过
                if key_text != "":
                    value_text = values_sc[index].replace("\"", "\\\"").replace("\n", "\\n").replace("\\n\"", "\"").lstrip()
                    map_sc[key_text] = value_text
                    sc_new_value = "\"%s%s\" = \"%s\"" % (key_prefix,key_text, value_text) # 加入 map_key_value_en
                    file_object.write(sc_new_value + ";\n")
                    file_sc.write(sc_new_value + ";\n")
                        

    # 写入文件 TC
    filename_tc = sheet_name + '_tc.strings'
    with open("tc.strings", 'a+') as file_tc:
        with open(filename_tc, 'w') as file_object:
            # file_object.write("/* %s */\n" % (sheet_name))
            print("filename_tc: %s" % filename_tc)
            for index in range(len(keys)):
                key_text = keys[index]
                # 如果KEY内容为空(""),将被跳过
                if key_text != "":
                    value_text = values_tc[index].replace("\"", "\\\"").replace("\n", "\\n").replace("\\n\"", "\"").lstrip()
                    tc_new_value = "\"%s%s\" = \"%s\"" % (key_prefix,key_text, value_text) # 加入 map_key_value_en
                    file_object.write(tc_new_value + ";\n")
                    file_tc.write(tc_new_value + ";\n")
                    map_tc[key_text] = value_text

# 写入文件 - 去重后
with open("deduplication_en.strings", 'w') as file_en:
    for key in keys_order_list:
        file_en.write("\"%s\" = \"%s\";\n" % (key, map_en[key]) )
with open("deduplication_sc.strings", 'w') as file_sc:
    for key in keys_order_list:
        file_sc.write("\"%s\" = \"%s\";\n" % (key, map_sc[key]) )
with open("deduplication_tc.strings", 'w') as file_tc:
    for key in keys_order_list:
        file_tc.write("\"%s\" = \"%s\";\n" % (key, map_tc[key]) )

print("SUCCESSSSSSSSSS!\n")
