# Tạo vòng lập lấy thông tin

# Tạo vòng lập lấy thông tin

import datetime
import os
import time
from tkinter.filedialog import askdirectory
import xlwings as xw

# sheet hiện hành
sht = xw.sheets.active

# lệnh hiện lên cửa sổ chọn thư mục. Lấy đườn dẫn thư mục gán vào biến
path = askdirectory(title="Chọn thư mục")

# lấy đường dẫn đến thư mục hiện hành để làm việc
os.chdir(path)

# lấy đường dẫn hiện hành, làm thư mục hiện hành làm việc
thumuc_hienhanh = os.getcwd()

#lấy tất cả tên file trong thư mục cho vào biến list
list_file = os.listdir()

# đưa giá trị vào ô A3
sht.range("A3").value = "Thư mục:" + thumuc_hienhanh

# số lượng file trong thư mục vào ô A4
sht.range("A4").value = "Số file:" + str(len(list_file))
row = 7
for name_file in list_file:
    # lấy đường dẫn đến file cần lấy thông tin
    path = os.path.join(thumuc_hienhanh,name_file)

    # dấu thời gian sửa đổi tập tin của một tập tin
    m_time = os.path.getmtime(path)  # ra một dãy số cần chuyển đổi dãy số này
    # ngày chỉnh sửa cuối cùng
    dt_m = datetime.datetime.fromtimestamp(m_time).strftime("%d/%m/%Y - %H:%M:%S")


    # lấy thời gian truy cập cuối cùng của tập tin
    a_time = os.path.getatime(path)  # ra một dãy số cần chuyển đổi để lấy thông tin
    # chuyển đổi thông tin ngày giờ
    dt_a = time.strftime("%d/%m/%Y - %H:%M:%S", time.localtime(a_time))

    # ngày giờ tạo tập tin
    c_time = os.path.getctime(path)
    # chuyển đổi định dạng ngày tạo file
    dt_c = time.strftime("%d/%m/%Y - %H:%M:%S", time.localtime(c_time))



    # Lấy thông dung lượng file MB ()
    size = os.path.getsize(path)
    size_new = round(size / 1024, 2)


    sht.range("A" + str(row)).value = name_file
    sht.range("b" + str(row)).value = thumuc_hienhanh
    sht.range("d" + str(row)).value = dt_m
    sht.range("e" + str(row)).value = dt_a
    sht.range("f" + str(row)).value = dt_c
    sht.range("g" + str(row)).value = size_new
    sht.range("H" + str(row)).value = "=HYPERLINK({0},{1})".format('"'+ path + '"','"Open"')
    sht.range("I" + str(row)).value = name_file

    row = row+1
