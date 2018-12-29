
import xlrd
import os
import numpy as np
import matplotlib.pyplot as plt


CWD = os.getcwd()
FILES_PATH = CWD + "\\files\\"
fname = FILES_PATH + 'input.xls'

file = xlrd.open_workbook(fname)
sheet_names = file.sheet_names()
sheet = file.sheet_by_name(sheet_names[0])

y = sheet.col_values(0)
x = sheet.col_values(1)


reg = np.polyfit(x,y,3)
ry = np.polyval(reg,x)

plt.plot(x,y,'r.',label='X-Y')
plt.plot(x,ry,label='f(x)')

plt.xlabel("X")
plt.ylabel("Y")

plt.show()
print(reg)
