import numpy as np
import pandas as pd
import sys

pd.options.mode.chained_assignment = None

def doProcess(INPATH,OUTPATH):

    CIOINPUT = pd.read_excel("CIO_Input.xlsx", sheet_name='CIO',skiprows = 1,usecols = "B:M")
    Dist = pd.read_excel("Distancekm.xlsm", sheet_name='Sheet1',usecols = "A:D")
    CIOINPUT['SOURCE']=CIOINPUT['EUtranCellFDD'].str[4:10]
    CIOINPUT['TARGET-ENB']=CIOINPUT['EUtranCellRelation'].str[5:10]
    Concat1 = CIOINPUT['SOURCE'] + CIOINPUT["TARGET-ENB"].map(str)
    CIOINPUT['CONCAT'] = Concat1

    df = pd.merge(CIOINPUT,Dist,left_on='CONCAT',right_on='CON',how='inner')

    df.drop(['CONCAT', 'CON', 'Unnamed: 1', 'Unnamed: 2'], axis = 1,inplace=True)
    df['Note'] = np.where((df.DIS == 0) & (df.Execfail > 0), 'CIO will be 0',
	np.where((df.DIS > 6000) & (df.Execfail > 0), 'CIO Reduce -6',
	np.where((df.DIS < 3000) & df.EarlyLate.str.contains('Early'), 'CIO Reduce -2',
	np.where((df.DIS < 3000) & df.EarlyLate.str.contains('Late'), 'CIO Increase +2',
	np.where((df.DIS < 6000) & df.EarlyLate.str.contains('Early'), 'Reduce -3 to -5',
	np.where((df.DIS < 6000) & df.EarlyLate.str.contains('Late'), 'Increase -3 to -5', ''))))))

    df.drop_duplicates()

    writer = pd.ExcelWriter('Output.xlsx')
    df.to_excel(writer, sheet_name='output')
    writer.save()

#doProcess(INPATH = r"C:\Users\earnhaz\Documents\TML\automation\HOSR\\", OUTPATH = r"C:\Users\earnhaz\Documents\TML\automation\HOSR\Output\\")
if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("Please provide input and output path")
        sys.exit()
    INPATH = sys.argv[1]
    OUTPATH = sys.argv[2]
    doProcess(INPATH, OUTPATH)