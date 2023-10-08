from win32com.client.gencache import EnsureDispatch
from win32com.client import constants

#测试失败

Excel = EnsureDispatch("Excel.Application") # 打开Excel程序
f = r"G:/MyCode/python2023/holdlist/933.xlsx"

wb = Excel.Workbooks.Open(f) # 打开Excel工作簿
print(wb)
print(help(wb))
sht = wb.Sheets("Sheet1") # 指定工作表
sht.PrintOut() # 打印工作表
wb.Close(constants.xlDoNotSaveChanges) # （不保存）关闭工作簿

Excel.Quit() # 退出Excel程序