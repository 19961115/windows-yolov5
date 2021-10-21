"""
@Author:范志光
@Date：2021/10/18
"""
import xlwt
#创建表格并赋值
def create_excel(file_name):
    data = ["windows-close", "windows-open", "windows-half-open"]
    workbook = xlwt.Workbook(encoding='utf-8')
    ws = workbook.add_sheet("windows")
    ws.write(0, 0, label=data[0])
    ws.write(0, 1, label=data[1])
    ws.write(0, 2, label=data[2])
    ws.write(0,3,"图片名称")
    ws.write(0,4,"窗户总数")
    ws.write(0,5,"% windows-close")
    ws.write(0,6,"% windows-open")
    ws.write(0,7,"% windows-half-open")

    workbook.save(file_name)

#设置某一行全为0的函数
def do_all_zero(i,file_name):
    import xlrd
    from xlutils.copy import copy
    # 打开想要更改的excel文件
    old_excel = xlrd.open_workbook(file_name, formatting_info=True)
    # 将操作文件对象拷贝，变成可写的workbook对象
    new_excel = copy(old_excel)
    # 获得第一个sheet的对象
    ws = new_excel.get_sheet(0)
    # 写入数据
    ws.write(i, 0, 0)
    ws.write(i,1,0)
    ws.write(i,2,0)
    ws.write(i, 3, 0)
    ws.write(i, 4, 0)
    ws.write(i, 5, 0)
    ws.write(i, 6, 0)
    ws.write(i, 7, 0)
    new_excel.save(file_name)

#修改表格数据函数，包括某行某列，包括数据大小，包括文件名称
def reator(i,j,num,file_name):
    import xlrd
    from xlutils.copy import copy
    # 打开想要更改的excel文件
    old_excel = xlrd.open_workbook(file_name, formatting_info=True)
    # 将操作文件对象拷贝，变成可写的workbook对象
    new_excel = copy(old_excel)
    # 获得第一个sheet的对象
    ws = new_excel.get_sheet(0)
    # 写入数据
    ws.write(i, j, num)
    new_excel.save(file_name)

#处理excel表格数据
import os
def myrename(path):
    file_list=os.listdir(path)
    i=100
    for fi in file_list:
        old_name=os.path.join(path,fi)
        new_name=os.path.join(path,str(i)+".jpg")
        os.rename(old_name,new_name)
        i+=1
#myrename("C:\\Users\\20170\\PycharmProjects\\students-inclassroom-monitoring-main\\exp\\")

#降低文件夹图片的分辨率
def decline_pr():
    import os
    from PIL import Image
    import glob

    img_path = glob.glob("C:/Users/20170/PycharmProjects/students-inclassroom-monitoring-main/exp/*.jpg")
    path_save = "C:/Users/20170/PycharmProjects/students-inclassroom-monitoring-main/exp1"
    for file in img_path:
        name = os.path.join(path_save, file)
        im = Image.open(file)
        im.thumbnail((1000, 1000))
        print(im.format, im.size, im.mode)
        im.save(name)
decline_pr()