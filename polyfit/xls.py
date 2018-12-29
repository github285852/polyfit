from xlwt import Workbook
import xlrd




"""
FUNCTION:
    从excel中读入数据，
INPUT:
    fname,文件名,
    start,开始坐标，(x0,y0) x,列，y，行
    end,(x1,y1),
    col_r , True，一列一列读，False，一行一行读
    sheetname,表格名称，默认读第一个表格
RETURN:
    numpy.array数据
"""
def ImportNumpyDataFromXls(fname,start,end,col_r = True,sheetname = ""):
    file = xlrd.open_workbook(fname)
    sheet_names = file.sheet_names()
    sheet = file.sheet_by_name(sheet_names[0])
    #print(sheet.cell(1,2).value)
    end_x = end[0]
    end_y = end[1]
    end_x = end_x + 1
    end_y = end_y + 1
    start_x = start[0]
    start_y = start[1]
    if col_r == True : #每列读
        data = np.zeros((end_x-start_x,end_y-start_y))
        for col in range(0,end_x - start_x):#列
            for row in range(0,end_y - start_y):
                #                               y            x
                data[col][row] = sheet.cell(row+start_y,col+start_x).value
    else:#每行每行读
        data = np.zeros((end_y - start_y,end_x-start_x))
        for row in range(0,end_y - start_y):
            for col in range(0,end_x - start_x):#列
                data[row][col] = sheet.cell(row+start_y,col+start_x).value
    return data



"""
向excel写数据，一行一行写
INPUT:pos,(x,y)x,列，y 行
sheet,要写的表格
col_w = Ture,一列一列写，False一行一行写
"""
def WriteExcel(pos,data,sheet,col_w=True):
    x,y = pos
    if col_w == False:
        for list1 in data:
            for i in range(0,len(list1)):
                sheet.write(y,x+i,list1[i])
            y += 1
        return y - pos[1]
    else:
        for list1 in data:
            for i in range(0,len(list1)):
                sheet.write(y+i,x,list1[i])
            x += 1
        return x - pos[1]


