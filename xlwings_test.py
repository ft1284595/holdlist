#https://blog.csdn.net/m0_46388544/article/details/123074260?ops_request_misc=&request_id=&biz_id=102&utm_term=python%E6%89%93%E5%8D%B0Excel%E6%96%87%E6%A1%A3&utm_medium=distribute.pc_search_result.none-task-blog-2allsobaiduweb~default-1-123074260.142v19pc_rank_34,157v15new_3&spm=1018.2226.3001.4187%20%E2%80%94%E2%80%94%E2%80%94%E2%80%94%E2%80%94%E2%80%94%E2%80%94%E2%80%94%E2%80%94%E2%80%94%E2%80%94%E2%80%94%E2%80%94%E2%80%94%E2%80%94%E2%80%94%20%E7%89%88%E6%9D%83%E5%A3%B0%E6%98%8E%EF%BC%9A%E6%9C%AC%E6%96%87%E4%B8%BACSDN%E5%8D%9A%E4%B8%BB%E3%80%8C%E6%9D%9C%E6%9D%9C123%E3%80%8D%E7%9A%84%E5%8E%9F%E5%88%9B%E6%96%87%E7%AB%A0%EF%BC%8C%E9%81%B5%E5%BE%AACC%204.0%20BY-SA%E7%89%88%E6%9D%83%E5%8D%8F%E8%AE%AE%EF%BC%8C%E8%BD%AC%E8%BD%BD%E8%AF%B7%E9%99%84%E4%B8%8A%E5%8E%9F%E6%96%87%E5%87%BA%E5%A4%84%E9%93%BE%E6%8E%A5%E5%8F%8A%E6%9C%AC%E5%A3%B0%E6%98%8E%E3%80%82%20%E5%8E%9F%E6%96%87%E9%93%BE%E6%8E%A5%EF%BC%9Ahttps://blog.csdn.net/weixin_48160417/article/details/125395557


# 打印一个工作簿中的所有工作表
import xlwings as xw
app = xw.App(visible=False, add_book=False)
workbook = app.books.open('各月销售数量表.xlsx')
workbook.api.PrintOut(Copies=2, ActivePrinter='DESKTOP-HP01', Collate=True) # 打印工作簿中的所有工作表，这里指定打印份数为两份，打印机为“DESKTOP-HP01”
workbook.close()
app.quit()

# 打印一个工作簿中的一个工作表
import xlwings as xw
app = xw.App(visible=False, add_book=False)
workbook = app.books.open('各月销售数量表.xlsx')
worksheet = workbook.sheets['1月']
worksheet.api.PrintOut(Copies=2, ActivePrinter='DESKTOP-HP01', Collate=True)
workbook.close()
app.quit()

# 打印多个工作簿
from pathlib import Path
import xlwings as xw
folder_path = Path('各地区销售数量')
file_list = folder_path.glob('*.xls*')
app = xw.App(visible=False, add_book=False)
for i in file_list:
    workbook = app.books.open(i)
    workbook.api.PrintOut(Copies=1, ActivePrinter='DESKTOP-HP01', Collate=True)
    workbook.close()
app.quit()

# 打印多个工作簿中的同名工作表
from pathlib import Path
import xlwings as xw
folder_path = Path('各地区销售数量')
file_list = folder_path.glob('*.xls*')
app = xw.App(visible=False, add_book=False)
for i in file_list:
    workbook = app.books.open(i)
    worksheet = workbook.sheets['销售数量']
    worksheet.api.PrintOut(Copies=1, ActivePrinter='DESKTOP-HP01', Collate=True)
    workbook.close()
app.quit()

# 打印工作表的指定单元格区域
import xlwings as xw
app = xw.App(visible=False, add_book=False)
workbook = app.books.open('销售表.xlsx')
worksheet = workbook.sheets['总表']
area = worksheet.range('A1:I10')
area.api.PrintOut(Copies=1, ActivePrinter='DESKTOP-HP01', Collate=True)
workbook.close()
app.quit()

# 按指定的缩放比例打印工作表
import xlwings as xw
app = xw.App(visible=False, add_book=False)
workbook = app.books.open('销售表.xlsx')
worksheet = workbook.sheets['总表']
worksheet.api.PageSetup.Zoom = 80 # 按工作表原始大小的80%进行打印
worksheet.api.PrintOut(Copies=1, ActivePrinter='DESKTOP-HP01', Collate=True)
workbook.close()
app.quit()

# 在纸张的居中位置打印工作表
import xlwings as xw
app = xw.App(visible=False, add_book=False)
workbook = app.books.open('各月销售数量表.xlsx')
worksheet = workbook.sheets['1月']
worksheet.api.PageSetup.CenterHorizontally = True # 调整工作表在纸张上的水平位置
worksheet.api.PageSetup.CenterVertically = True # 调整工作表在纸张上的垂直位置
worksheet.api.PrintOut(Copies=1, ActivePrinter='DESKTOP-HP01', Collate=True)
workbook.close()
app.quit()

# 打印工作表时打印行号和列号
import xlwings as xw
app = xw.App(visible=False, add_book=False)
workbook = app.books.open('各月销售数量表.xlsx')
worksheet = workbook.sheets['1月']
worksheet.api.PageSetup.PrintHeadings = True
worksheet.api.PrintOut(Copies=1, ActivePrinter='DESKTOP-HP01', Collate=True)
workbook.close()
app.quit()

# 重复打印工作表的标题行
import xlwings as xw
app = xw.App(visible=False, add_book=False)
workbook = app.books.open('销售表.xlsx')
worksheet = workbook.sheets['总表']
worksheet.api.PageSetup.PrintTitleRows = '$1:$1' # 将工作表第一行设置为要重复打印的标题行
worksheet.api.PageSetup.Zoom = 55 # 按工作表原始大小的55%进行打印
worksheet.api.PrintOut(Copies=1, ActivePrinter='DESKTOP-HP01', Collate=True)
workbook.close()
app.quit()
