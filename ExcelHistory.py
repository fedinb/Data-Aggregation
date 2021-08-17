import pandas as pd
import numpy as np
from datetime import date, time

# These variables are needed for the log file which is created at the end of this script.

start_time = time.time()
today = date.today()
d = today.strftime("%B %d, %Y")


dfPivot = pd.read_csv('C:\\Python_Input_Files\\DIY\\History.csv')

# Here we clean possible NaN values. They may appear due to technical errors on WMS side.

dfPivot = dfPivot.dropna()

# Below we create the Demand column.

dfPivot['Demand'] = dfPivot['Shipped'] + dfPivot['OOS']
    
# String format for columns with Product codes and GLN codes

dfPivot[['GLN_Code','EAN_Code','Base_Unit_Code']] = dfPivot[['GLN_Code','EAN_Code','Base_Unit_Code']].astype(str)

# The negative values in columns with quantities are filtered out. 

Movement_Cols = ['Stock','Shipped','Transit','OOS','Demand']

dfPivot[Movement_Cols] = dfPivot[dfPivot[Movement_Cols] >= 0][Movement_Cols]

dfPivot[Movement_Cols] = dfPivot[Movement_Cols].astype(int)


# Names of DCs according to their GLN-codes.

GLN_Names = {'4080398539419': 'Toronto',
              '4080398340519': 'Calgary',
              '4080398955559': 'Montreal',
              '4080398400515': 'Vancouver',
              '4080398955085': 'Edmonton'}

# Dataframe DF now has a column with DCs' names.

dfPivot['GLN_Name'] = dfPivot['GLN_Code'].map(GLN_Names)

# Columns' order is changed for dataframe.

dfPivot = dfPivot[['Movement_Date','GLN_Code','GLN_Name','EAN_Code','Base_Unit_Code','Stock','Transit','Shipped','OOS','Demand']]

# XLSXWriter splits whole dataframe into sheets according to their DCs' name.

GLN_Names = dfPivot.GLN_Name.unique().tolist()

all_sheets = GLN_Names + ['All DCs']

writer = pd.ExcelWriter('C:\\Python_Output_Files\\DIY\\History.xlsx',
                        engine='xlsxwriter')

dfPivot.to_excel(writer,sheet_name='All DCs',index=False)


for GLN in GLN_Names:
    xslxdf = dfPivot.loc[dfPivot.GLN_Name==GLN]
    xslxdf.to_excel(writer, sheet_name=GLN, index=False)
    
# Optimized column width is set for the sheets.    
    
for column in dfPivot:
    column_width = max(dfPivot[column].astype(str).map(len).max(), len(column))
    col_idx = dfPivot.columns.get_loc(column)
    for sheet in all_sheets:
        writer.sheets[sheet].set_column(col_idx, col_idx, column_width+2)

writer.save()

# File with logs is created. It will store basic date for all script's iterations.
    
def ScriptLog(filename):
    date = today.strftime("%B %d, %Y")
    duration = str(np.round(time.time() - start_time,decimals=1))
    tab_1 = ' ' * (20 - len(date))
    tab_2 = ' ' * (10 - len(duration))
    l = open(filename,"a+")
    l.write('Date: '+date+tab_1+'Seconds to run: '+duration+tab_2+'Total number of rows: '+str(dfPivot.shape[0])+'\n')
    l.close
    
ScriptLog('C:\\Python_Output_Files\\DIY\\History_Logs.txt')



