import xlwings as xw    
import datetime
import os

def listFiles(path):
    '''
        根据指定的路径,遍历所有后缀名是xlsx的文件,这个方法只遍历了当前目录,并没有迭代遍历当前路径下的子目录
    '''
    files_list = []
    for file in os.listdir(path):
        print(file)
        #print(type(file))
        #~开头的文件是临时文件
        if(file.startswith('~')):
            continue
        if os.path.splitext(file)[1] == '.xlsx':
            files_list.append(file)
    print("---------valid print file start---------------")
    print(files_list)
    print("---------valid print file end-----------------")
    return files_list

def main():

    for filename in listFiles('.'):

        #默认打印1份
        copy_num = 1
        #app = xw.App(visible=True, add_book=False)
        app = xw.App(visible=False, add_book=False)
        workbook = app.books.open(filename)
        worksheet = workbook.sheets['zx-pivot-table']
        #print('last_cell row=',worksheet.used_range.last_cell.row)
        #数据透视表的头部和尾部统计,一共有4行
        if worksheet.used_range.last_cell.row > 14:     #如果表格的数据(具体要找的包裹数量)超过10行,打印2份
            copy_num = 2
        if worksheet.used_range.last_cell.row > 24:     #如果表格的数据超(具体要找的包裹数量)过20行,打印3份
            copy_num = 3
        

        #在左边的页眉打印当前日期
        worksheet.api.PageSetup.LeftHeader = datetime.datetime.now().strftime("%Y-%m-%d")
        worksheet.api.PrintOut(Copies=copy_num, ActivePrinter="Brother MFC-L2717DW Printer (Copy 1)",Collate=True)

        app.quit()


main()