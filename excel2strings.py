import xlrd
import os

# log重复KEY信息,保留不去重的strings和去重后的strings

########################################### 初始化定义 - 开始
# 指定Excel文件名, (别忘加后缀)
excel_file_name = "nsl1.4.xlsx"
# 指定sheets名字s
sheet_names = ["Reset secondary password", "Reset password", "Error handling", "Device optional page"]
sheet_key_prefixs = ["", "", "", ""]

#### 指定关键行号

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

# 为了保证key的导出顺序,用list来记录key的出现顺序,如果有key重复出现,将会移除前一个key. 
# (另外,经测试,从map遍历key-value导出也是同一个结果,以防万一乱序)
# 去重后的有序 key list
keys_ordered_list = [] 
# 去重后的key-value,全部sheet遍历完成后再写入
map_en = {}
map_sc = {}
map_tc = {}

# excel对象
book = xlrd.open_workbook(excel_file_name)

def trim_key(text):
    return text.lstrip().rstrip()
def trim_value(text):
    return text.lstrip().rstrip().replace("\"", "\\\"").replace("\n", "\\n").replace("\\n\"", "\"")

if os.path.exists("all_keys") == False:
    os.system("mkdir all_keys & mkdir all_keys/en all_keys/sc all_keys/tc")
if os.path.exists("de_key") == False:
    os.system("mkdir de_key")
if os.path.exists("log") == False:
    os.system("mkdir log")    

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
    with open("./all_keys/en.strings", 'w') as file_en:
        filename_en = "./all_keys/en/%s_en.strings" % (sheet_name)
        with open(filename_en, 'w') as file_object:
            for index in range(len(keys)):
                key_text = trim_key(keys[index])
                if key_text != "": 
                    value_text = trim_value(values_en[index])
                    map_new_record = "sheet:%s:%s" % (sheet_name, (index + rowIndexForStartKeyValue + 1)) # 定位到原文件中sheet+行号
                    map_en[key_text] = value_text
                    if keys_ordered_list.count(key_text) != 0:
                        keys_ordered_list.remove(key_text)
                    keys_ordered_list.append(key_text)
                    # 处理提示
                    if map_record.get(key_text) is None:  # 无重复
                        map_record[key_text] = map_new_record # 加入 map_record
                    else:
                        map_old_record = map_record.get(key_text)
                        map_record[key_text] = "%s ; %s" % (map_old_record,map_new_record)
                        with open("./log/log.strings", 'w') as log:
                            log.write("------------------------------\n")
                            log.write("Duplicate Key:%s\n" % key_text)
                            log.write("%s\n\n" % map_record[key_text])
                    en_new_value = "\"%s%s\" = \"%s\"" % (key_prefix,key_text, value_text) # 加入 map_key_value_en
                    file_object.write(en_new_value + ";\n")
                    file_en.write(en_new_value + ";\n")
                        
    # 写入文件 SC
    with open("./all_keys/sc.strings", 'w') as file_sc:
        filename_sc = "./all_keys/sc/%s_sc.strings" % (sheet_name)
        with open(filename_sc, 'w') as file_object:
            for index in range(len(keys)):
                key_text = trim_key(keys[index])
                if key_text != "":
                    value_text = trim_value(values_sc[index])
                    sc_new_value = "\"%s%s\" = \"%s\"" % (key_prefix,key_text, value_text) # 加入 map_key_value_sc
                    file_object.write(sc_new_value + ";\n")
                    file_sc.write(sc_new_value + ";\n")
                    map_sc[key_text] = value_text
                        

    # 写入文件 TC
    with open("./all_keys/tc.strings", 'w') as file_tc:
        filename_tc = "./all_keys/tc/%s_tc.strings" % (sheet_name)
        with open(filename_tc, 'w') as file_object:
            for index in range(len(keys)):
                key_text = trim_key(keys[index])
                if key_text != "":
                    value_text = trim_value(values_tc[index])
                    tc_new_value = "\"%s%s\" = \"%s\"" % (key_prefix,key_text, value_text) # 加入 map_key_value_tc
                    file_object.write(tc_new_value + ";\n")
                    file_tc.write(tc_new_value + ";\n")
                    map_tc[key_text] = value_text

# 写入文件 - 去重后
with open("./de_key/de_key_en.strings", 'w') as file_en:
    for key in keys_ordered_list:
        file_en.write("\"%s\" = \"%s\";\n" % (key, map_en[key]) )
with open("./de_key/de_sc.strings", 'w') as file_sc:
    for key in keys_ordered_list:
        file_sc.write("\"%s\" = \"%s\";\n" % (key, map_sc[key]) )
with open("./de_key/de_tc.strings", 'w') as file_tc:
    for key in keys_ordered_list:
        file_tc.write("\"%s\" = \"%s\";\n" % (key, map_tc[key]) )

print("Complete!!!")
