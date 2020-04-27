import xlrd
from xlutils.copy import copy
import BachKhoa
import NgoaiNgu
import CNTT_TT
import PH_Kontum
import SuPham
import SP_KyThuat
import NCDT_VietAnh
import YDuoc
import GDTC
import KinhTe

QDTT = '3043/QĐ-ĐHĐN ngày 18/9/2019'
maTruong = 'DDI'
namTT = '2019'
trinhDo = 'Đại học'
hinhThuc = 'Chính quy'
doiTuong = 'THPT'
phuongThuc = 'Xét Tuyển'
nguoiLam_QD = 'NGUYỄN ĐĂNG HUY'


file_location = "I DSTT kem QD 2019 - Dot bo sung lan 1.xls"

wb = xlrd.open_workbook(file_location)
sheet = wb.sheet_by_index(0)

rb = xlrd.open_workbook('DSTT kem QD - Moi.xls')
sheet_rows = rb.sheet_by_index(0)
numOfRows = sheet_rows.nrows

wt = copy(rb)
sheet1 = wt.get_sheet(0)


for rows in range(sheet.nrows):
    if 7 < rows:
        # if sheet.cell_value(rows, 0) == '':
        #     tenNganh = sheet.cell_value(rows, 1).replace('Ngành: ', '').strip()
        #     maNganh = CNTT_TT.maNganh(tenNganh)
        # elif type(sheet.cell_value(rows, 0) is str and ):
        #     break
        if type(sheet.cell_value(rows, 0)) is float:
            for cols in range(4):
                value = sheet.cell_value(rows, cols + 1)
                if cols == 0:
                    sheet1.write(numOfRows, 9, value)
                elif cols == 1:
                    arr_ten = value.split(' ')
                    ten = arr_ten.pop()
                    tenDem = ' '.join(arr_ten)
                    sheet1.write(numOfRows, 1, tenDem)
                    sheet1.write(numOfRows, 2, ten)
                elif cols == 2:
                    ngay = value.split('/')[0]
                    thang = value.split('/')[1]
                    nam = value.split('/')[2]
                    sheet1.write(numOfRows, 3, ngay)
                    sheet1.write(numOfRows, 4, thang)
                    sheet1.write(numOfRows, 5, nam)
                else:
                    sheet1.write(numOfRows, 7, value)
            sheet1.write(numOfRows, 10, namTT)
            sheet1.write(numOfRows, 11, trinhDo)
            sheet1.write(numOfRows, 12, hinhThuc)
            sheet1.write(numOfRows, 13, doiTuong)
            sheet1.write(numOfRows, 14, phuongThuc)
            sheet1.write(numOfRows, 15, maTruong)
            sheet1.write(numOfRows, 16, maNganh)
            sheet1.write(numOfRows, 17, tenNganh)
            sheet1.write(numOfRows, 18, '-')
            sheet1.write(numOfRows, 22, QDTT)
            sheet1.write(numOfRows, 23, nguoiLam_QD)
            numOfRows += 1
        else:
            if sheet.cell_value(rows, 0) == '':
                tenNganh = sheet.cell_value(rows, 1).replace('Ngành: ', '').strip()
                maNganh = CNTT_TT.maNganh(tenNganh)
            else:
                break


# sheet1.write(2, 0, 'sample 2')
#
wt.save('DSTT kem QD - Moi.xls')
