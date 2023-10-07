import pandas as pd
import numpy as np



def main():
    df = pd.read_excel("./369.xlsx", dtype=str, engine='openpyxl')
    df.head()

    pt = pd.pivot_table(df, index=["MAWB","Container No.","HAWB"], values=["CBP Status"], aggfunc=["count"], margins=True)
    

    print(pt)
    print(type(pt))
    
    #column_widths = (pt.columns.to_series().apply(lambda x : len(x.encode('utf-8'))).values)

    max_widths = (pt.astype(str).applymap(lambda x : len(x.encode('utf-8'))).agg(max).values)

    #print(column_widths)
    print(max_widths)

    #https://cloud.tencent.com/developer/article/1770494



    pt.to_excel("../369-pivot.xlsx")
    """
    with pd.ExcelWriter("933.xlsx") as writer:
        pt.to_excel("../933-pivot.xlsx", sheet_name="zx-pivot-table")
        worksheet = writer.sheets['zx-pivot-table']
        for i in range(1,3):
            worksheet.set_column(i, i, 20)

        writer.save()
    """













main()