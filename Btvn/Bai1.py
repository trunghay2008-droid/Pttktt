import timeit 
from openpyxl import Workbook

a = 2 # co dinh a = 2
x = 2 # co dinh x = 2

def cach1():
    kq = 0
    for i in range(n):
        kq += a * (x ** i)
    return kq

def cach2():
    kq = 2
    for i in range(n - 1):
        kq = 2 + x * kq
    return kq

wb = Workbook()
ws = wb.active
ws.title = "Ket qua"
ws.append(["n", "Thoi gian cach 1", "Thoi gian cach 2", "Nhan xet"])

print("Dang chay...")
for n in [1000, 10000, 100000]:
    nx = ""
    time1 = timeit.timeit(lambda: cach1(), number=1)
    time2 = timeit.timeit(lambda: cach2(), number=1)
    if (time1 > time2): nx = f"cach1 cham hon cach2 {time1 - time2} giay"
    elif (time1 < time2): nx = f"cach1 nhanh hon cach2 {time2 - time1} giay"
    ws.append([n, time1, time2, nx])
print("Xong")

wb.save("ketqua.xlsx")
    
