import xlrd
import math

Workbook = xlrd.open_workbook('demo19.xlsx')
sheet1 = Workbook.sheet_by_name('Sheet1')

myCell = sheet1.cell(0, 0)


# Case 1 : calc the average value.

shibor_list = []
libor_list = []
hibor_list = []
# print(sheet1.nrows)
# print(sheet1.cell(1,1).value)
for i in range(1, sheet1.nrows):
    shibor_list.append(sheet1.cell(i, 1).value)
    libor_list.append(sheet1.cell(i, 2).value)
    hibor_list.append(sheet1.cell(i, 3).value)


avg = lambda x: sum(x) / len(x)
avg_shibor = avg(shibor_list)
avg_hibor = avg(hibor_list)
avg_libor = avg(libor_list)


# Case 2: calc the variability.
def r_sigma(lists, a):
    sigma = math.sqrt(sum(map(lambda x: (x - a) ** 2, lists)) / (len(lists) - 1))
    return sigma


print(r_sigma(shibor_list, avg_shibor))

# Case 3:
print(avg_shibor, avg_hibor, avg_libor)
print(r_sigma(shibor_list, avg_shibor), r_sigma(hibor_list, avg_hibor), r_sigma(libor_list, avg_libor))