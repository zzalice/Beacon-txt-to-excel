# coding=utf8
import xlwt

# 此檔用於協助公式計算
#
# excel檔會長這樣:
# 距離 | HTC和Sony四方位的rssi平均值(一共約240個rssi值的平均)
# 0.5  | 81
# 1	   | 82
# 1.5  | 78
# 2	   | 77
# 2.5  | 79
# 3	   | 79
# 3.5  | 76
# 4    | 71
# 4.6  | 65
# 5	   | 62
# 5.5  | 64
# 6	   | 72


open_file = [['./6_6/HTC/1_back_beacon/9_20170607 01_14_01.txt',
              './6_6/HTC/2_face_beacon/9_20170607 01_28_02.txt',
                './6_6/HTC/3_face_door/9_20170607 01_38_11.txt',
                './6_6/HTC/4_face_wall/9_20170607 01_47_59.txt',
             './6_6/sony/1_back_beacon/9_20170607 01_13_57.txt',
             './6_6/sony/2_face_beacon/9_20170607 01_27_50.txt',
               './6_6/sony/3_face_door/9_20170607 01_38_09.txt',
               './6_6/sony/4_face_wall/9_20170607 01_47_56.txt'
              ],
             ['./6_6/HTC/1_back_beacon/9_20170607 01_18_02.txt',
              './6_6/HTC/2_face_beacon/9_20170607 01_28_43.txt',
                './6_6/HTC/3_face_door/9_20170607 01_38_53.txt',
                './6_6/HTC/4_face_wall/9_20170607 01_48_46.txt',
             './6_6/sony/1_back_beacon/9_20170607 01_18_01.txt',
             './6_6/sony/2_face_beacon/9_20170607 01_28_41.txt',
               './6_6/sony/3_face_door/9_20170607 01_38_51.txt',
               './6_6/sony/4_face_wall/9_20170607 01_48_43.txt'
              ],
             ######################################1
             ['./6_6/HTC/1_back_beacon/9_20170607 01_18_48.txt',
              './6_6/HTC/2_face_beacon/9_20170607 01_29_28.txt',
                './6_6/HTC/3_face_door/9_20170607 01_39_36.txt',
                './6_6/HTC/4_face_wall/9_20170607 01_49_29.txt',
             './6_6/sony/1_back_beacon/9_20170607 01_18_46.txt',
             './6_6/sony/2_face_beacon/9_20170607 01_29_25.txt',
               './6_6/sony/3_face_door/9_20170607 01_39_33.txt',
               './6_6/sony/4_face_wall/9_20170607 01_49_27.txt'
              ],
             ['./6_6/HTC/1_back_beacon/9_20170607 01_19_32.txt',
              './6_6/HTC/2_face_beacon/9_20170607 01_31_28.txt',
                './6_6/HTC/3_face_door/9_20170607 01_40_18.txt',
                './6_6/HTC/4_face_wall/9_20170607 01_50_12.txt',
             './6_6/sony/1_back_beacon/9_20170607 01_19_37.txt',
             './6_6/sony/2_face_beacon/9_20170607 01_31_27.txt',
               './6_6/sony/3_face_door/9_20170607 01_40_15.txt',
               './6_6/sony/4_face_wall/9_20170607 01_50_09.txt'
              ],
             ######################################2
             ['./6_6/HTC/1_back_beacon/9_20170607 01_20_21.txt',
              './6_6/HTC/2_face_beacon/9_20170607 01_32_20.txt',
                './6_6/HTC/3_face_door/9_20170607 01_41_00.txt',
                './6_6/HTC/4_face_wall/9_20170607 01_50_55.txt',
             './6_6/sony/1_back_beacon/9_20170607 01_20_19.txt',
             './6_6/sony/2_face_beacon/9_20170607 01_32_18.txt',
               './6_6/sony/3_face_door/9_20170607 01_40_57.txt',
               './6_6/sony/4_face_wall/9_20170607 01_50_52.txt'
              ],
             ['./6_6/HTC/1_back_beacon/9_20170607 01_21_03.txt',
              './6_6/HTC/2_face_beacon/9_20170607 01_33_05.txt',
                './6_6/HTC/3_face_door/9_20170607 01_41_41.txt',
                './6_6/HTC/4_face_wall/9_20170607 01_51_46.txt',
             './6_6/sony/1_back_beacon/9_20170607 01_21_00.txt',
             './6_6/sony/2_face_beacon/9_20170607 01_33_02.txt',
               './6_6/sony/3_face_door/9_20170607 01_41_39.txt',
               './6_6/sony/4_face_wall/9_20170607 01_51_43.txt'
              ],
             ######################################3
             ['./6_6/HTC/1_back_beacon/9_20170607 01_21_47.txt',
              './6_6/HTC/2_face_beacon/9_20170607 01_33_05.txt',
                './6_6/HTC/3_face_door/9_20170607 01_42_23.txt',
                './6_6/HTC/4_face_wall/9_20170607 01_52_26.txt',
             './6_6/sony/1_back_beacon/9_20170607 01_21_44.txt',
             './6_6/sony/2_face_beacon/9_20170607 01_33_44.txt',
               './6_6/sony/3_face_door/9_20170607 01_42_20.txt',
               './6_6/sony/4_face_wall/9_20170607 01_52_26.txt'
              ],
             ['./6_6/HTC/1_back_beacon/9_20170607 01_22_28.txt',
              './6_6/HTC/2_face_beacon/9_20170607 01_33_05.txt',
                './6_6/HTC/3_face_door/9_20170607 01_43_48.txt',
                './6_6/HTC/4_face_wall/9_20170607 01_53_12.txt',
             './6_6/sony/1_back_beacon/9_20170607 01_22_26.txt',
             './6_6/sony/2_face_beacon/9_20170607 01_34_26.txt',
               './6_6/sony/3_face_door/9_20170607 01_43_46.txt',
               './6_6/sony/4_face_wall/9_20170607 01_53_11.txt'
              ],
             ######################################4
             ['./6_6/HTC/1_back_beacon/9_20170607 01_23_11.txt',
              './6_6/HTC/2_face_beacon/9_20170607 01_35_10.txt',
                './6_6/HTC/3_face_door/9_20170607 01_44_29.txt',
                './6_6/HTC/4_face_wall/9_20170607 01_54_07.txt',
             './6_6/sony/1_back_beacon/9_20170607 01_23_09.txt',
             './6_6/sony/2_face_beacon/9_20170607 01_35_08.txt',
               './6_6/sony/3_face_door/9_20170607 01_44_27.txt',
               './6_6/sony/4_face_wall/9_20170607 01_53_59.txt'
              ],
             ['./6_6/HTC/1_back_beacon/9_20170607 01_23_53.txt',
              './6_6/HTC/2_face_beacon/9_20170607 01_35_51.txt',
                './6_6/HTC/3_face_door/9_20170607 01_45_11.txt',
                './6_6/HTC/4_face_wall/9_20170607 01_54_54.txt',
             './6_6/sony/1_back_beacon/9_20170607 01_23_51.txt',
             './6_6/sony/2_face_beacon/9_20170607 01_35_48.txt',
               './6_6/sony/3_face_door/9_20170607 01_45_09.txt',
               './6_6/sony/4_face_wall/9_20170607 01_54_51.txt'
              ],
             ######################################5
             ['./6_6/HTC/1_back_beacon/9_20170607 01_26_11.txt',
              './6_6/HTC/2_face_beacon/9_20170607 01_36_32.txt',
                './6_6/HTC/3_face_door/9_20170607 01_45_52.txt',
                './6_6/HTC/4_face_wall/9_20170607 01_55_41.txt',
             './6_6/sony/1_back_beacon/9_20170607 01_26_08.txt',
             './6_6/sony/2_face_beacon/9_20170607 01_36_30.txt',
               './6_6/sony/3_face_door/9_20170607 01_45_51.txt',
               './6_6/sony/4_face_wall/9_20170607 01_55_39.txt'
              ],
             ['./6_6/HTC/1_back_beacon/9_20170607 01_26_57.txt',
              './6_6/HTC/2_face_beacon/9_20170607 01_37_14.txt',
                './6_6/HTC/3_face_door/9_20170607 01_46_46.txt',
                './6_6/HTC/4_face_wall/9_20170607 01_56_24.txt',
             './6_6/sony/1_back_beacon/9_20170607 01_26_54.txt',
             './6_6/sony/2_face_beacon/9_20170607 01_37_12.txt',
               './6_6/sony/3_face_door/9_20170607 01_46_50.txt',
               './6_6/sony/4_face_wall/9_20170607 01_56_23.txt'
              ]
                   ######################################6

             ]
store_to = './6_6/6_6_indoor_beacon9.xls'
new_sheet = "6_6_indoor_beacon9"

# Create workbook and sheet
workbook = xlwt.Workbook(encoding='utf8')

# Add new sheet
sheet = workbook.add_sheet(new_sheet, cell_overwrite_ok=True)


sheet.write(0, 0, u"距離")
sheet.write(0, 1, u"HTC和Sony四方位的rssi平均值(一共約240個rssi值的平均)")

# 填距離欄位
d = 0.5
for c in range(1, 13):
    sheet.write(c, 0, d)
    d += 0.5


# 存excel檔
workbook.save(store_to)

# 算兩支手機、在同一個距離、收同一個beacon的rssi值的平均值(方位不管，都一起加入平均)
excel_row_count = 0
excel_col_count = 1
for distance in range(0, len(open_file)):                   # 有len(open_file)個距離
    excel_row_count += 1
    for files in range(0, len(open_file[distance])):        # 有len(open_file[distance])個要平均
        # 讀txt
        f = open_file[distance][files].decode('utf8')       # 用utf8解碼
        with open(f) as f:
            data = f.readlines()                            # 讀檔

        count_rssi = 0                                      # 數檔案裡有幾個rssi
        rssi = 0                                            # rssi加總
        hundred = 1                                         # 紀錄要往後讀幾個字元

        print "dis: ", distance, "  files: ", files

        for j in range(4, len(data[0]), 4):
            if data[0][0] == '':                                    # 沒有數據
                print("error!!!!!!!")
            elif int(data[0][j-3+hundred:j-1+hundred])<20:          # 如果rssi是三位數
                print("error!!!!!!!")
                rssi += int(data[0][j-3+hundred:j+hundred])
                                   # 只存三位數，不存空白與負號
                                               # 沒有減一，因為要多讀一個字元
                hundred += 1
                count_rssi += 1
            else:
                rssi += int(data[0][j-3+hundred:j-1+hundred])
                                   # 只存二位數，不存空白與負號
                count_rssi += 1

        rssi /= count_rssi                                  # 算平均
        sheet.write(excel_row_count, excel_col_count, rssi)
        workbook.save(store_to)
