# lấy thông tin các file trong thư mục do người dùng lựa chọn

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
