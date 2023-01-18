import datetime
import os
import time

# Đường dẫn 1 file trong thư mục
path = r"H:\My Drive\HOCPYTHON excel\mo dau\SD\so_diem_khoi_10_mon_dia_li.xls"

# dấu thời gian sửa đổi tập tin của một tập tin
m_time = os.path.getatime(path)
# đổi thông tin ra ngày giờ
dt_m = datetime.datetime.fromtimestamp(m_time).strftime("%d/%m/%Y, %H:%M:%S")
print(dt_m)

# lấy thời gian truy cập cuối cùng của tập tin
lan_cuoi = os.path.getatime(path)
print(lan_cuoi)
# chuyển đổi thông tin ngày giờ
thoigian_lan_cuoi = time.strftime("%d/%m/%Y, %H:%M:%S", time.localtime(lan_cuoi))
print(thoigian_lan_cuoi)
