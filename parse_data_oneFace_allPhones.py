# coding=utf8
import xlwt


# 此檔用於處理存了無數個rssi值的txt檔
# 可用來輸出一個excel檔，裡面放了n台手機、面對同一個方位、在同一個點上、接收同一個rssi的 n*30個rssi值
# excel檔會長這樣:
# HTC | SonyXA      -->n台手機
# 91  | 68
# 84  | 67
# 84  | 68
# 89  | 69
# ......            -->30個值，如果要更多，可以改52行第程式碼
# 73  | 89
# 85  | 80
# 83  | 76


# 欲開啟的txt檔路徑
open_file = ['./6_6/HTC/3_face_door/9_20170607 01_38_11.txt',
             './6_6/sony/3_face_door/9_20170607 01_38_09.txt'
            ]
# 輸出的excel檔的儲存路徑
store_to = './6_6/for_graph/faceDoor/50/6_6_beacon14_faceDoor_dis50.xls'
# 新建立的excel檔裡的sheet名稱
new_sheet = "6_6_beacon14_faceDoor_dis50"

# Create workbook and sheet
workbook = xlwt.Workbook(encoding='utf8')

# Add new sheet
sheet = workbook.add_sheet(new_sheet, cell_overwrite_ok=True)

# 寫入
sheet.write(0, 0, u"HTC")
sheet.write(0, 1, u"SonyXA")

# 存excel檔
workbook.save(store_to)


for t in range(0, len(open_file)):          # 開啟t個txt檔

    f = open_file[t].decode('utf8')         # 用utf8解碼
    with open(f) as f:
        data = f.readlines()                # 讀檔

    excel_row_count = 0
    excel_col_count = t
    hundred = 1                             # 紀錄

    for j in range(4, 124, 4):              # 只讀取30個rssi值，用四的倍數是因為通常rssi都是兩位數，
                                            # 所以會是空白+負號+兩個數字，共四個字元
        excel_row_count += 1
        if data[0][j] == '':
            print("error")
        elif int(data[0][j-3+hundred:j-1+hundred]) < 20:        # 萬一rssi是三位數
            print("error")
            sheet.write(excel_row_count, excel_col_count, int(data[0][j-3+hundred:j+hundred]))
                                                                    # 只存二位數，不存空白與負號
                                                                                # 沒有減一，因為要多讀一個字元
            workbook.save(store_to)
            hundred += 1                    # 紀錄之後都要往後讀一個字元
        else:
            sheet.write(excel_row_count, excel_col_count, int(data[0][j-3+hundred:j-1+hundred]))
                                                                    # 只存二位數，不存空白與負號
            workbook.save(store_to)

