import datetime
import os
import time
from tkinter.filedialog import askdirectory

# thiết lập chọn thư mục
path = askdirectory(title="Chọn thư mục")

# thư mục làm việc
os.chdir(path)

# Lấy đường dẫn hiện tại làm thư mục hiện hành
pff = os.getcwd()

# lấy tất cả các file trong thư mục
list_file = os.listdir()
print(list_file)
print(len(list_file))



"""
# Đường dẫn 1 file trong thư mục
path = r"H:\My Drive\HOCPYTHON excel\mo dau\SD\so_diem_khoi_10_mon_dia_li.xls"

# lấy thông tin từ file trong thư mục ngày giờ [Cách 1: Dùng hàm datetime]
m_time = os.path.getatime(path)

# đổi thông tin ra ngày giờ
dt_m = datetime.datetime.fromtimestamp(m_time).strftime("%d/%m/%Y, %H:%M:%S")
print(dt_m)

# lấy thời gian chỉnh sửa lần cuối cùng [cách 2: dùng hàm os và time]
lan_cuoi = os.path.getatime(path)
print(lan_cuoi)
# chuyển đổi thông tin ngày giờ
thoigian_lan_cuoi = time.strftime("%d/%m/%Y, %H:%M:%S", time.localtime(lan_cuoi))
print(thoigian_lan_cuoi)
"""
