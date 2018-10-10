#!/usr/bin/python3
import xlrd
import xlwt
import pymongo

# --------------------文件路径-------------------------------
week = 36
file = 'EH DRR analysis by wk '+str(week)+' 2018 - WMUS For Meeting.xlsx'
file2 = 'sales report week '+str(week)+'.xlsx'


# ----------------------------------------------------------


def readxlxs(filename):  # 读取excel
    try:
        data = xlrd.open_workbook(filename)
        return data
    except Exception:
        print(filename)
        print('open xlsx error!')


def save_in_VSNdata():
    VSNdata.delete_many({})
    table_ALL_DRR = readxlxs(file).sheet_by_name(u'All Items Weekly DRR')
    Col_1 = table_ALL_DRR.col_values(1)
    Col_4 = table_ALL_DRR.col_values(4)
    for index in range(len(Col_1)):
        if Col_1[index] != "" and Col_4[index] != "Vendor Stk Nbr" and Col_4[index] != "":
            if '\n新' in Col_4[index]:
                Col_4[index] = Col_4[index].split('\n新')[1]
            if '(grey fabric)' in Col_4[index]:
                Col_4[index] = Col_4[index].split(' (grey fabric)')[0]

            temp = {"_id": index + 1, "VSN": Col_4[index], "WK"+str(week-2)+"P": int(table_ALL_DRR.col_values(156 + (week - 1) * 3)[index]),
                    "WK"+str(week-2)+"C": int(table_ALL_DRR.col_values(
                        157 + (week - 1) * 3)[index]), "WK"+str(week-1)+"P": int(table_ALL_DRR.col_values(159 + (week - 1) * 3)[index]),
                    "WK"+str(week-1)+"C": int(table_ALL_DRR.col_values(160 + (week - 1) * 3)[index])}
            VSNdata.insert_one(temp)
            del temp
        else:
            temp = {"_id": index + 1, "VSN": Col_4[index], "WK"+str(week-2)+"P": 0,
                    "WK"+str(week-2)+"C": 0, "WK"+str(week-1)+"P": 0,"WK"+str(week-1)+"C": 0}
            VSNdata.insert_one(temp)
            del temp


def save_in_WSRdata():  # 根据VSNdata中的内容来搜寻创建数据表
    WSRdata.delete_many({})
    table_raw_data = readxlxs(file2).sheet_by_name(u'raw date')
    i = 20
    j = 1
    while i < table_raw_data.nrows:
        if table_raw_data.col_values(7)[i] == 'POS Qty' or table_raw_data.col_values(7)[i] == 'Cust Def Qty':
            WSRdata.insert_one({'_id': j, 'VSN': table_raw_data.col_values(3)[i], 'Type': table_raw_data.col_values(
                7)[i], '2018'+str(week-1): int(table_raw_data.col_values(62 + week)[i]),
                '2018'+str(week): int(table_raw_data.col_values(62 + week + 1)[i])})
        print(i)
        i = i + 1
        j = j + 1
    print('sucessful')


def write_new_xlsx():
    wk = xlwt.Workbook(encoding='utf-8')
    ws = wk.add_sheet('my_sheet', cell_overwrite_ok=True)
    VSN_List = VSNdata.find()
    for i in range(VSNdata.count()):
        cur = VSNdata.find()[i]['VSN']
        ws.write(i, 0, cur)
        findresult = WSRdata.find({'VSN': cur})
        matchresult = VSNdata.find({'VSN':cur})[0]['WK'+str(week-1)+'P']
        for perfindresult in range(0, findresult.count(), 2):
            if (matchresult == findresult[perfindresult]['2018'+str(week-1)]):
                ws.write(i, 1, findresult[perfindresult]['2018'+str(week-1)])
                ws.write(i, 2, findresult[perfindresult + 1]['2018'+str(week-1)])
                ws.write(i, 3, findresult[perfindresult]['2018'+str(week)])
                ws.write(i, 4, findresult[perfindresult + 1]['2018'+str(week)])
    wk.save('Excel_test.xls')

# ----------------Main-------------------------------
# ----------------连接数据库--------------------------
myclient = pymongo.MongoClient("mongodb://localhost:27017/")
mydb = myclient["autoDRR"]
VSNdata = mydb["VSNdata"]
WSRdata = mydb["WSRdata"]

# 读取表格内容，存入数据库
save_in_VSNdata()
save_in_WSRdata()
write_new_xlsx()
