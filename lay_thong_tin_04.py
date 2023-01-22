# Tạo vòng lập lấy thông tin

# Tạo vòng lập lấy thông tin

import datetime
from itertools import count
import os
from tkinter.filedialog import askdirectory
import xlwings as xw

sht = xw.sheets.active

path = askdirectory(title="Chọn thư mục")
os.chdir(path)

#đường dẫn hiện hành
dd_hienhanh = os.getcwd()

# lấy danh sách file trong thư mục hiện hành cho vào list
list_file = os.listdir()

# ghi giá trị vào ô A3 của sheet hiện hành đang mở
sht.range("A3").value = "Thư mục chứa file: " + dd_hienhanh

# ghi số lượng file trong thư mục vào ô A4
sht.range("A4").value = "Số file trong thư mục:" + str(len(list_file))

sht.range("A6").value = "Tên file"
sht.range("B6").value = "Thư mục chứa file"
sht.range("C6").value = "Thời gian chỉnh sửa file"
sht.range("D6").value = "Thời gian chỉnh sửa file lần cuối"
sht.range("E6").value = "Thời gian khởi tạo file"
sht.range("F6").value = "Dung lượng file (KB)"
sht.range("G6").value = "Mở xem file"
sht.range("H6").value = "Tên file cần chỉnh sửa"


dong = 7
# Tạo vòng lập lấy thông số các file trong thư mục
for name_file in list_file:
    # ghép đường dẫn hiện hành với từng tên file trong thư mục
    path_2 = os.path.join(dd_hienhanh,name_file)

    # thời gian sửa đổi file
    m_time = os.path.getmtime(path_2)
    dt_m = datetime.datetime.fromtimestamp(m_time).strftime("%d/%m/%Y - %H:%M:%S")

    # thời gian truy cập file cuối cùng
    a_time = os.path.getatime(path_2)
    dt_a = datetime.datetime.fromtimestamp(a_time).strftime("%d/%m/%Y - %H:%M:%S")

    # thời gian khởi tạo tập tin
    c_time = os.path.getctime(path_2)
    dt_c = datetime.datetime.fromtimestamp(c_time).strftime("%d/%m/%Y - %H:%M:%S")

    # lấy dung lượng file KB
    size = os.path.getsize(path_2)
    size_file = round(size/1024,2)

    # đưa các giá trị thông tin file vào các ô
    sht.range("A" + str(dong)).value = name_file
    sht.range("B" + str(dong)).value = dd_hienhanh
    sht.range("C" + str(dong)).value = dt_m
    sht.range("D" + str(dong)).value = dt_a
    sht.range("E" + str(dong)).value = dt_c
    sht.range("F" + str(dong)).value = size_file
    sht.range("G" + str(dong)).value = "=HYPERLINK({0},{1})".format('"'+path_2+'"','"Mở xem file"')
    sht.range("H" + str(dong)).value = name_file

    dong = dong + 1

