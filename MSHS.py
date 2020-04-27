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

QDTT = '3352/QĐ-ĐHĐN ngày 01/10/2018'
maTruong = 'DDP'
namTT = '2019'
trinhDo = 'Đại học'
hinhThuc = 'Chính quy'
doiTuong = 'THPT'
phuongThuc = 'Xét Tuyển'
nguoiLam_QD = 'NGUYỄN ĐĂNG HUY'


file_location = "DSTT hoc ba BS dot 1 kem QD.xls"

wb = xlrd.open_workbook(file_location)
sheet = wb.sheet_by_index(1)

rb = xlrd.open_workbook('DSTT kem QD - Moi.xls')
sheet_rows = rb.sheet_by_index(0)
numOfRows = sheet_rows.nrows

wt = copy(rb)
sheet1 = wt.get_sheet(0)


for rows in range(sheet.nrows):
    if 7 < rows:
        if type(sheet.cell_value(rows, 0)) is float:
            for cols in range(4):
                value = sheet.cell_value(rows, cols + 1)
                if cols == 3:
                    # SBD/MSHS
                    sbd = sheet.cell_value(rows, 7)
                    sheet1.write(numOfRows, 9, sbd)
                elif cols == 0:
                    # Họ Tên
                    ho_ten = sheet.cell_value(rows, 1).upper()
                    arr_ten = ho_ten.split(' ')
                    ten = arr_ten.pop()
                    tenDem = ' '.join(arr_ten)
                    sheet1.write(numOfRows, 1, tenDem)
                    sheet1.write(numOfRows, 2, ten)
                elif cols == 1:
                    # Ngày sinh
                    ngaySinh = sheet.cell_value(rows, 2)
                    ngay = ngaySinh.split('/')[0]
                    thang = ngaySinh.split('/')[1]
                    nam = ngaySinh.split('/')[2]
                    sheet1.write(numOfRows, 3, ngay)
                    sheet1.write(numOfRows, 4, thang)
                    sheet1.write(numOfRows, 5, nam)
                else:
                    # Giới tính
                    gioiTinh = sheet.cell_value(rows, 3)
                    sheet1.write(numOfRows, 7, gioiTinh)
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
                maNganh = PH_Kontum.maNganh(tenNganh)
            else:
                break


# sheet1.write(2, 0, 'sample 2')
#
wt.save('DSTT kem QD - Moi.xls')
