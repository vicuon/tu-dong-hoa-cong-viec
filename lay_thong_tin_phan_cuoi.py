import datetime
from itertools import count
import os
from tkinter.filedialog import askdirectory
import xlwings as xw
from tkinter import messagebox
import filetype

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

    # lấy thông tin phần mềm tạo file
    kind = filetype.guess(path_2)
    split_tup = os.path.splitext(path_2)

    # đưa các giá trị thông tin file vào các ô
    sht.range("A" + str(dong)).value = name_file
    sht.range("B" + str(dong)).value = dd_hienhanh
    sht.range("C" + str(dong)).value = dt_m
    sht.range("D" + str(dong)).value = dt_a
    sht.range("E" + str(dong)).value = dt_c
    sht.range("F" + str(dong)).value = size_file
    sht.range("G" + str(dong)).value = "=HYPERLINK({0},{1})".format('"'+path_2+'"','"Mở xem file"')
    #sht.range("H" + str(dong)).value = name_file
    #sht.range("I" + str(dong)).value = kind.mime if kind != None else "'không biết'"
    sht.range("I" + str(dong)).value = split_tup[1] # [1]: lấy phần đuôi file; [0]: lấy tên và đường dẫn file

    dong = dong + 1

"""giai đoạn 2
# Xuất hiện thông báo yes/no. 
# Nếu chọn yes thực hiện đổi tên file
# chọn no không thực hiện
"""
# Hiện thông báo chọn Yes/No
res = messagebox.askquestion("Thông báo","Bạn muốn thay đổi tên tập tin không")

# Tìm số dòng cuối trong bản tính 
lr = sht.range("A"+str(sht.cells.last_cell.row)).end("up").row
print(lr)

#Vùng tên fiel cũ
range_old_name = xw.Range("A7:A"+str(lr)).value
# vùng tên file mới
range_new_name = xw.Range("H7:H"+str(lr)).value

# Nếu nhấn nút chọn yes => muốn đổi tên file
if res == "yes":
    for i in range(len(range_new_name)):
        os.rename(range_old_name[i], range_new_name[i])

"""
Cần biết thông tin phần mềm  tạo file
cần cài đặt thư viện
pip install filetype
hoặc dùng "split_tup = os.path.splitext(path_2)"
"""
