from openpyxl import load_workbook

wb = load_workbook(filename = 'string.xlsx')

key_list = [] # 主要为记录顺序
map_en = {}
map_tc = {}
map_sc = {}

def trim_key(text):
    return text.strip()
def trim_value(text):
    if text is None:
        return ""
    return text.strip().replace("\"", "\\\"").replace("\n", "\\n").replace("\\n\"", "\"")

with open("strings.log", 'w') as file_log:
    for sheet in wb:
        file_log.write("---------------------------sheet_name:%s----------------------\n" % sheet.title)
        for row in sheet.iter_rows(min_row=2,min_col=1,max_col=4, values_only=True):
            # blank line: 4 cell all None, break this inner layer loop.
            if (row[0] is None) and (row[1] is None) and (row[2] is None) and (row[3] is None):
                break
            # blank cell: 1 or 2, or 3 cell is None, write to log
            elif (row[0] is None) or (row[1] is None) or (row[2] is None) or (row[3] is None):
                file_log.write("skip line[%s]:%s,%s,%s,%s\n" % (sheet.title,row[0],row[1],row[2],row[3]))
                continue
            # OK line: write to xx.strings
            else:
                key = trim_key(row[0])
                value_en = trim_value(row[1])
                value_tc = trim_value(row[2])
                value_sc = trim_value(row[3])

                if key_list.count(key) != 0:
                    key_list.remove(key)
                key_list.append(key)
                map_en[key] = value_en
                map_tc[key] = value_tc
                map_sc[key] = value_sc

with open("en.strings", 'w') as file_en:
    with open("sc.strings", 'w') as file_sc:
        with open("tc.strings", 'w') as file_tc:
            for key in key_list:
                value_en = map_en[key]
                value_tc = map_tc[key]
                value_sc = map_sc[key]
                file_en.write("\"%s\" = \"%s\";\n" % (key, value_en))
                file_sc.write("\"%s\" = \"%s\";\n" % (key, value_sc))
                file_tc.write("\"%s\" = \"%s\";\n" % (key, value_tc))

print("===============================END====================================\n")
