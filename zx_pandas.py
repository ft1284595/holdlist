import pandas as pd
import os

'''
holdlist汇总预处理, 从excel文件中取出'MAWB','HAWB','Container No.'这3列的内容,然后把这3列的内容保存成一个新的excel文件
保存的位置是当前路径的上一级目录
'''

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
    #print(111111111111111)
    print(files_list)
    #print(111111111111111)
    return files_list

def main():
    #df = pd.read_excel("933.xlsx")
    #df = pd.read_excel("369.xlsx")
    
    """
    print(df['Container No.'].values)
    print("=============")
    print(df['HAWB'].values)
    print("*************")
    #print(df[df['On Hold']!='HOLD']['MAWB'].values[0])
    print(df['MAWB'].values)
    """

    awb_list = []
    container_list = []
    hawb_list = []
    count_list = []

    #for holdlist in ['297.xlsx']:
    for holdlist in listFiles('.'):
    #for holdlist in listFiles('C:\\Users\\tyler\\Downloads\\0918'):
        df = pd.read_excel(holdlist, dtype=str)
        #print(df['HAWB'].values)
        #print('attentation before sort')
        #print(df['Container No.'].values)
        df.sort_values(by=['Container No.'], inplace=True)
        #print('after sort')
        #print(df['Container No.'].values)

        for i, item in enumerate(df['MAWB'].values):
            if i == 0:
                awb_list.append(item)
            else:
                awb_list.append('')

        for item in df['Container No.'].values:
            container_list.append(item)

        
        for i, item in enumerate(df['HAWB'].values):
            #print('***************************')
            #print(item)
            #print('---------------------------')
            #print(str(item))
            #print('***************************')
            hawb_list.append(str(item))
            if i < len(df['HAWB'].values) - 1:
                count_list.append('')
            elif i == len(df['HAWB'].values) - 1:
                count_list.append(len(df['HAWB'].values))
            else:
                raise Exception("emmmmmmmmmmmmmmmmmmm, something wrong, check HAWB column")

    
    
    #print(awb_list)
    #print(container_list)
    #print(hawb_list)
    #print(count_list)

    content = pd.DataFrame({'AWB':awb_list, 'Container No.':container_list, "HAWB": hawb_list, "Count":count_list})
    #writer = pd.ExcelWriter("zx.xlsx", mode='a',if_sheet_exists='overlay')
    writer = pd.ExcelWriter("../zx-hold-list.xlsx")
    content.to_excel(writer, index=False)
    writer.close()





main()