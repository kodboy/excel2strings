import xlrd

# 去重

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

book = xlrd.open_workbook(excel_file_name)

map_record = {} # Check en only
keys_order_list = [] # 为了保证key的导出顺序,用list来记录key的出现顺序,如果有key重复出现,将会移除前一个key. 另外,经测试,从map遍历key-value导出也是同一个结果.
map_en = {}
map_sc = {}
map_tc = {}

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
    
    # 加入map_en
    filename_en = sheet_name + '_en.strings'
    for index in range(len(keys)):
        key_text = keys[index]
        if key_text != "": # 如果KEY内容为空(""),将被跳过
            value_text = values_en[index].replace("\"", "\\\"").replace("\n", "\\n").replace("\\n\"", "\"").lstrip()
            map_en[key_text] = value_text
            if keys_order_list.count(key_text) != 0:
                keys_order_list.remove(key_text)
            keys_order_list.append(key_text)
                        
    # 加入map_sc
    filename_sc = sheet_name + '_sc.strings'
    for index in range(len(keys)):
        key_text = keys[index]
        if key_text != "": # 如果KEY内容为空(""),将被跳过
            value_text = values_sc[index].replace("\"", "\\\"").replace("\n", "\\n").replace("\\n\"", "\"").lstrip()
            map_sc[key_text] = value_text

    # 加入map_tc
    filename_tc = sheet_name + '_tc.strings'
    for index in range(len(keys)):
        key_text = keys[index]
        if key_text != "": # 如果KEY内容为空(""),将被跳过
            value_text = values_sc[index].replace("\"", "\\\"").replace("\n", "\\n").replace("\\n\"", "\"").lstrip()
            map_tc[key_text] = value_text
# 写入文件
with open("en.strings", 'w') as file_en:
    for key in keys_order_list:
        file_en.write("\"%s\" = \"%s\";\n" % (key, map_en[key]) )
with open("sc.strings", 'w') as file_sc:
    for key in keys_order_list:
        file_sc.write("\"%s\" = \"%s\";\n" % (key, map_sc[key]) )
with open("tc2.strings", 'w') as file_tc:
    for key in keys_order_list:
        file_tc.write("\"%s\" = \"%s\";\n" % (key, map_tc[key]) )

print(len(keys_order_list))
print("导出成功,已去重!")
