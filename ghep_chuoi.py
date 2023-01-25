"""
Ghép các chuổi lại bằng 03 cách:
1. Ghép bằng phép cộng
2. Ghép bằng F-string
3. Ghép bằng format
"""

name = "Nguyễn Việt Cường"
camp_name = "GV THPT"
job = "Giáo viên"

# Toán tử công đề nối các chuổi lại với nhau
info_1 = "+ Chuổi 1: xin chào, tôi là " + name + \
    " tôi đang làm việc tại " + camp_name +\
        "với vị trí là " + job
print(info_1)

# định dạng chuổi sử dụng F-string
info_2 = f"+ Chuổi 2: xin chào, tôi là {name}\
tôi đang làm việc tại {camp_name}\
với vị trí là {job}"
print(info_2)

# phương thức format, các biến gán từ trái sang phải
info_3 = "+ Chuổi 3: xin chào, tôi là {}\
tôi đang làm việc tại {}\
với vị trí là {}".format(name,camp_name,job)
print(info_3)
