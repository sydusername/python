import math

import win32com.client

# 1、填写配置信息
point_xy_path = r'C:\Users\xxxx\Desktop\xxxx.xls'

# 2、运行excel程序
xlapp = win32com.client.DispatchEx('Excel.Application')
xlapp.Visible = False  # true打开excel程序界面
xlapp.DisplayAlerts = False  # 禁止弹窗-不显示警告信息

# 3、打开excel，选择工作簿
xlbook1 = xlapp.Workbooks.Open(point_xy_path, ReadOnly=False)  # 打开文件
xsheet1 = xlbook1.Worksheets[0]  # 工作簿，首个工作簿序号是0
maxRow1 = xsheet1.UsedRange.Rows.Count
print('maxRow1', maxRow1)

# 4、加载excel数据
original_date = []
for i in range(2, maxRow1 + 1):
    original_date.append([float(xsheet1.Cells(i, 1).value), float(xsheet1.Cells(i, 2).value)])
print("original_date")

# 5、循环排序
modify_date = []
row = 0
distance = 0
while True:
    try:
        date_x = original_date[row][0]
        date_y = original_date[row][1]
        modify_date.append([date_x, date_y, distance])
        del original_date[row]
    except:
        break

    # 寻找最近项
    distance = 10000
    for i in range(len(original_date)):
        distance_temporary = math.sqrt((date_x - original_date[i][0]) ** 2 + (date_y - original_date[i][1]) ** 2)

        if distance > distance_temporary:
            distance = distance_temporary
            row = i
print("modify_date")

# 6、写入excel
xsheet1.Range(xsheet1.Cells(2, 1), xsheet1.Cells(maxRow1, 3)).Value = modify_date
print('save')

# 7、另存excel
xlbook1.save
new_excel_path = point_xy_path.split(".")
xlbook1.SaveAs(new_excel_path[0] + "_modify." + new_excel_path[1])
xlbook1.Close()
xlapp.Quit()
print("finsh")
