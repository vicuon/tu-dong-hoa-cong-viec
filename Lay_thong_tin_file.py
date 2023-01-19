import time
import datetime
import os

path = r"D:\lap trinh python\IA\SD\so_diem_khoi_10_mon_dia_li.xls"

# file modification timestamp of a file
# dấu thời gian sửa đổi tệp của tệp
m_time = os.path.getmtime(path)
# convert timestamp into DateTime object
#chuyển đổi dấu thời gian thành đối tượng DateTime
dt_m = datetime.datetime.fromtimestamp(m_time).strftime("%d/%m/%Y-%H:%M:%S")
print('Modified on:', dt_m)

# file creation timestamp in float
# Thời gian tạo file
c_time = os.path.getctime(path)
# convert creation timestamp into DateTime object
dt_c = datetime.datetime.fromtimestamp(c_time).strftime("%d/%m/%Y-%H:%M:%S")
print('Created on:', dt_c)

# file accessed timestamp in float
# Thời gian truy cập file cuối cùng
a_time = os.path.getatime(path)
# convert creation timestamp into DateTime object
dt_a = datetime.datetime.fromtimestamp(a_time).strftime("%d/%m/%Y-%H:%M:%S")
print('Accessed on:', dt_a)

# get size of file
# lấy dung lượng file
size = os.path.getsize(path)
print("Size file: ", round(size/1024,2), "KB")
