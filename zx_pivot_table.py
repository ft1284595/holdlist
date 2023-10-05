import pandas as pd
import numpy as np



def main():
    df = pd.read_excel("./933.xlsx", dtype=str)
    df.head()

    pt = pd.pivot_table(df, index=["MAWB","Container No.","HAWB"], values=["CBP Status"], aggfunc=["count"], margins=True)
    

    print(pt)
    pt.to_excel("../933-pivot.xlsx")
    """
    with pd.ExcelWriter("933.xlsx") as writer:
        pt.to_excel("../933-pivot.xlsx", sheet_name="zx-pivot-table")
        worksheet = writer.sheets['zx-pivot-table']
        for i in range(1,3):
            worksheet.set_column(i, i, 20)

        writer.save()
    """













main()