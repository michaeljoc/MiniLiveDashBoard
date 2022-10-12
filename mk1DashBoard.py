from asyncio.windows_events import NULL
from textwrap import fill
from tkinter.ttk import Style
import pyodbc 
import openpyxl
import warnings
import os
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File

import pandas as pd
from tabulate import tabulate
import plotly.graph_objects as go


from dash import Dash, dcc, html, Input, Output, dash_table 
import plotly.express as px
import datetime
import plotly

import sched
import time
from datetime import date
import numpy as np
import math
from guppy import hpy


app = Dash(__name__)
warnings.simplefilter(action='ignore', category=UserWarning)
warnings.simplefilter(action='ignore', category=FutureWarning)

df = pd.read_csv("test.csv")
df2 = pd.read_csv("testLeave.csv")

def testFunction(n): #Basically returns time format
    if(math.isnan(n)):
        print("Nan")
    else:
        #print("Activated test function")
        hours = int(n)
        minutes = (n*60) % 60
        seconds = (n*3600) % 60

        strReturn = str("%d:%02d" % (hours, minutes))
        return strReturn

def splitDataFrame(df): #This is to split the rows into two
    if len(df) % 2 != 0:
        df = df.iloc[:-1, :]
    df1, df2 =  np.array_split(df, 2)
    return df1, df2

app.layout = html.Div(id = "colTest",   children=[
    
    
    html.Div(id='live-update-text'),
    
    #html.P(datetime.datetime.today().strftime('%d/%m/%Y'), style={}),
    #html.A(html.Button(id='update_clicks',children= 'Refresh Data'),href='/'),
    dash_table.DataTable(
        id='table1',
        columns=[{"name": i, "id": i} 
                 for i in df.columns],
        data=df.to_dict('records'),
        #style_cell=dict(textAlign='left'),
        style_cell={'textAlign':'left'},
        #style_header=dict(backgroundColor="paleturquoise"),
        style_header={"backgroundColor":"#F9D056", 'font-weight': 'bold', 'height': 'auto', 'whiteSpace': 'normal', 'lineHeight': '10px'},
        #style_data=dict(backgroundColor="lavender") ,
        style_data= {'lineHeight': '-1px', 'backgroundColor':"#DF9065", 'height': 'auto',},
        editable=True,
        sort_action='native',
        sort_by=[{'column_id': 'Job ID', 'direction': 'asc'}],
        style_table= {'width': 'auto',  'float': 'left', 'height': 'auto'},
        style_data_conditional=[
            {
                'if': 
                {
                    'filter_query': '{Remaining} < 0',
                    'column_id': 'Remaining'
                },
                'backgroundColor': '#FF0001',
                'color': 'white',
                'font-weight': 'bold'
            },
            {
                'if': {
                    'filter_query': '{Remaining} > 0 && {Remaining} < 2',
                    'column_id': 'Remaining'
                },
                'backgroundColor': '#E2D152',
                'color': 'black',
                'font-weight': 'bold'
            }
        ],
        
        
        
        
        
        
        
    ),
    dash_table.DataTable(
        id='table2',
        columns=[{"name": i, "id": i} 
                 for i in df2.columns],
        data=df2.to_dict('records2'),
        style_cell=dict(textAlign='left'),
        style_header={"backgroundColor":"paleturquoise", 'font-weight': 'bold', 'height': 'auto', 'whiteSpace': 'normal', 'lineHeight': '10px'},
        style_data=dict(backgroundColor="lavender"),
        editable=True,
        sort_action='native',
        style_table= {'width': 'auto',  'float': 'right'},
        sort_by=[{'column_id': 'Employee', 'direction': 'asc'}]
        
        
        
        
        
    ),
    
    
    
    dcc.Interval(id='interval-component',interval=1*10000, n_intervals=0 )
], )



#style={'width': '50%',  'display': 'inline-block'},
#style_table={'display':'inline-block'},
#It works

@app.callback(
    Output('table1', 'data'), 
    Input('interval-component', 'n_intervals'))
def update_metrics(n):
    conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=iaas-eci-br.comvision.solutions;'
                      'Database=M1_BS;'
                      'UID=REDACTED;'
                      'PWD=REDACTED')

    cursor = conn.cursor()

    clockedOn = "Select lmeEmployeeName as Employee_Name,FORMAT(lmpActualStartTime,'hh:mm') as Start_Time From Timecards inner join Employees on lmpEmployeeID=lmeEmployeeID Where lmpActive = 1 and lmeShopEmployee=1 and lmeTerminationDate is null and lmeHomeProductionDepartmentID = 'PROD' and lmpTimecardDate = CAST( GETDATE() AS Date) order by lmeEmployeeID"

    notClockedOn = "Select lmeEmployeeName as Not_Clocked_Employee_Name From Employees Where lmeShopEmployee=1 and lmeTerminationDate is null and lmeEmployeeID not in (select lmpEmployeeID from Timecards where lmpActive = 1) and lmeHomeProductionDepartmentID = 'PROD' order by lmeEmployeeID"    

    employeeWorkCentres = "select lmeEmployeeName, lmljobid, lmlWorkCenterID as Work_Centre, lmlProcessID as Process_ID, jmoEstimatedProductionHours, Actual_Hours, ((Remaining_Hours * 60) - ((DATEPART(HOUR,GETDATE()) * 60) + DATEPART(MINUTE,GETDATE()) - ((DATEPART(HOUR,LastClockOn) * 60) + DATEPART(MINUTE,LastClockOn))))/60 as minsPlusNonClockRemain, FORMAT(LastClockOn,'hh:mm') as LastClockOn from ( Select distinct lmpEmployeeID, lmlJobID, lmeEmployeeName, lmlWorkCenterID, lmlProcessID, jmoEstimatedProductionHours, sum(distinct jmoActualProductionHours + jmoActualSetupHours) as Actual_Hours,  sum(distinct jmoEstimatedProductionHours) - sum(distinct jmoActualProductionHours + jmoActualSetupHours) as Remaining_Hours, LastClockOn =  max(  lmlActualStartTime ), row_number() over ( partition by lmpEmployeeID order by max(lmlActualStartTime ) desc ) as row_num_reverse From Timecards Inner Join Employees on lmpEmployeeID=lmeEmployeeID Inner Join TimecardLines on LMLTIMECARDID = LMPTIMECARDID left join JobOperations on LMLJOBID = JMOJOBID and LMLJOBASSEMBLYID = JMOJOBASSEMBLYID and LMLJOBOPERATIONID = JMOJOBOPERATIONID Where lmpActive = 1 and lmeShopEmployee=1 and lmeTerminationDate is null and CAST( GETDATE() AS Date ) = lmpTimecardDate and lmlEmployeeID = lmeEmployeeID group by lmpEmployeeID, lmeEmployeeName , lmlWorkCenterID, lmlProcessID, lmlJobID, jmoEstimatedProductionHours ) timecards where row_num_reverse = 1 order by lmpEmployeeID"

    leavesToday = "select distinct lmeEmployeeName , lmpLeaveAccrualID from Timecards Inner Join employees on LMPEMPLOYEEID = LMEEMPLOYEEID where lmpTimecardDate = CAST( GETDATE() AS Date) and lmpLeaveAccrualID != '' and lmeShopEmployee=1 and lmeTerminationDate is null and lmeHomeProductionDepartmentID = 'PROD' "



    df = pd.read_sql(clockedOn, conn)
    df2 = pd.read_sql(notClockedOn, conn)
    df3 = pd.read_sql(employeeWorkCentres, conn)
    df4 = pd.read_csv('reason.csv')
    df5 = pd.read_sql(leavesToday, conn)

    frames = [df, df2]
    #result = pd.concat(frames,  axis=1) #Important to have axis=1
    df.to_csv('./quickTesting/dfOutput.csv', index=True)
    df3.to_csv('./quickTesting/df3Output.csv', index=True)
    result = pd.concat([df3.set_index('lmeEmployeeName'),df.set_index('Employee_Name')], axis=1, join='outer')
    #result = result.rename(columns={'Employee_Name':'Employees (Clocked On)','Start_Time':'Start Time', 'Not_Clocked_Employee_Name':'Employee Name (Not Clocked On)'})

    df5 = df5.drop_duplicates(subset=["lmeEmployeeName"], keep='first')

    df2 = pd.concat([df2.set_index('Not_Clocked_Employee_Name'),df5.set_index('lmeEmployeeName')], axis=1, join='outer') #Joins if existing reason



    strExtend = df2.index.values
    strReasons = df2['lmpLeaveAccrualID'].values

    lenDiff = len(result) -len(strExtend)
    lenDiff2 = len(result) -len(strReasons)

    current = strExtend.tolist()
    currentReasons = strReasons.tolist()
    for x in range(0, lenDiff):
        current.append(' ')
    for x in range(0, lenDiff2):
        currentReasons.append(' ')

    if(len(result) < len(current)):
        listResult = result.values.tolist()
        columns = [result.index.name] + [i for i in result.columns]
        rows = [[i for i in row] for row in result.itertuples()]
        #print(len(rows))
        for x in range(0, len(current) - len(listResult)):
            #print("loop")
            rows.append(' ')

        result = pd.DataFrame(rows)
    else:
        listResult = result.values.tolist()
        columns = [result.index.name] + [i for i in result.columns]
        rows = [[i for i in row] for row in result.itertuples()]
        result = pd.DataFrame(rows)

    for x in range(0, len(result)-len(currentReasons)):
        currentReasons.append(' ')
    #print(current)
    #result['Not_Clocked_On'] = current
    #result['Reason'] = currentReasons

    result = result.rename(columns={0:'Employees',1:'Job ID', 2:'Work Centre', 3:'Process ID',4:'Estimated', 5:'Actual', 6:'Remaining', 7:'Latest', 8:'Start', 9:'Start Time', 'Process_IDsss':'Process ID'})
    #print(result.columns)
    todaysDate = datetime.datetime.today().strftime('%d/%m/%Y')
    result.index.names = [todaysDate] #Renames the index, you cannot use result.rename sadly
    result.reset_index().rename(columns={'index': ''})
    result.index += 1 

    #print(result['Job ID'].duplicated())

    ids = result["Job ID"]
    workCentreDuplicates = result["Work Centre"]


    duplicates = result[ids.isin(ids[ids.duplicated()])].sort_values("Job ID")

    secondDuplicates = pd.concat(g for _, g in result.groupby("Job ID") if len(g) > 1) #Gets Job Number Duplicates 
    secondDuplicates.to_csv("testDuplicates1.csv") 

    thirdDuplicates = pd.concat(g for _, g in secondDuplicates.groupby("Work Centre") if len(g) > 1)#Gets Work Centre Duplicates 

    thirdDuplicates.to_csv("testDuplicates2.csv")

    fourthDuplicates = pd.concat(g for _, g in thirdDuplicates.groupby("Process ID") if len(g) > 1) #Gets Work Centre Duplicates 

    fourthDuplicates.to_csv("testDuplicates3.csv")

    fifthDuplicates = pd.concat(g for _, g in fourthDuplicates.groupby("Job ID") if len(g) > 1) #Runs another job number duplicates and filters them out
    fifthDuplicates.to_csv("testDuplicates4.csv") 

    #Now remove all sites as they do not get halved no matter how many man powers
    cleanDuplicates = fifthDuplicates.drop(fifthDuplicates[(fifthDuplicates['Work Centre'] == "SITE") | (fifthDuplicates['Work Centre'] == "SITEA") | (fifthDuplicates['Work Centre'] == "SITET")].index)

    cleanDuplicates.to_csv("testDuplicates5.csv")  #cleanDuplicates has the rows that need dividing based on how many job numbers

    jobNumberList = cleanDuplicates['Job ID'].tolist() #This contains the list, now need to get the dividers next to it (count then assign number next to id)
    #print(jobNumberList)

    my_dict = {i:jobNumberList.count(i) for i in jobNumberList}
    #print(my_dict)

    toDivide = result.loc[result['Job ID'].isin(my_dict)].index.values #This gets the Index rows that contains the Job ID
    #print(toDivide)

    #print(my_dict)

    jobIDGet = "12498-003-001"

    #for jobID, divisor in my_dict.items():
        #if jobID == jobIDGet:
            #print(divisor)
    #result.to_csv('./quickTesting/result1.csv', index=True)

    #element = result.at[toDivide[0], "Remaining"] = result.at[toDivide[0], "Remaining"] / 2
    #print(result.at[toDivide[0], "Remaining"])
    listRows = []
    for x in toDivide:
        #print(result.at[x, "Job ID"])
        for jobID, divisor in my_dict.items():
            if jobID == result.at[x, "Job ID"]:
                divider = divisor
        
        #print(divider)
        #print(result.at[x, "Actual"])
        #print(result.at[x, "Estimated"])
        #print(result.at[x, "Actual"])
        #print(result.at[x, "Remaining"])
        #print(result.at[x, "Remaining"] - (result.at[x, "Estimated"] - result.at[x, "Actual"]))
        timeSinceClock = (result.at[x, "Remaining"] - (result.at[x, "Estimated"] - result.at[x, "Actual"])) / divider #Basically gets the other half of halved timeonclock
        #print(timeSinceClock)
        result.at[x, "Remaining"] = result.at[x, "Remaining"] / divider + timeSinceClock
        result.at[x, "Estimated"] = result.at[x, "Estimated"] / divider
        #print(result.at[x, "Actual"])
        result.at[x, "Actual"] = result.at[x, "Actual"] / divider
        
        listRows.append(result.at[x, "Remaining"])

    #print(listRows)

    result = result.round(1)

    
    
    #result["Remaining"] = result["Remaining"].apply(lambda row: testFunction(row))
    #Above works if you want decimal hours to hours:minutes, be warned you cannot filter accordingly

    #print(returnedResult)
    
    #-------------------Experimental, delete if it doesn't work
    #test1, test2 = splitDataFrame(result)
    #print(test2)
    #result = pd.concat([test1, test2], axis=1)
    #-------------------Experimental, delete if it doesn't work
    #h = hpy()
    #print(h.heap())
    
    
    result.to_csv("test.csv", index=True) #change index to true/false if you want it or if you want it gone

    df = pd.read_csv("test.csv")
    
    
    return df.to_dict('records')



@app.callback(
    Output('table2', 'data'), 
    Input('interval-component', 'n_intervals'))
def update_nonClock(n):
    conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=iaas-eci-br.comvision.solutions;'
                      'Database=M1_BS;'
                      'UID=REDACTED;'
                      'PWD=REDACTED')

    cursor = conn.cursor()

    

    notClockedOn = "Select lmeEmployeeName as Not_Clocked_Employee_Name From Employees Where lmeShopEmployee=1 and lmeTerminationDate is null and lmeEmployeeID not in (select lmpEmployeeID from Timecards where lmpActive = 1) and lmeHomeProductionDepartmentID = 'PROD' order by lmeEmployeeID"    

    

    leavesToday = "select distinct lmeEmployeeName , lmpLeaveAccrualID from Timecards Inner Join employees on LMPEMPLOYEEID = LMEEMPLOYEEID where lmpTimecardDate = CAST( GETDATE() AS Date) and lmpLeaveAccrualID != '' and lmeShopEmployee=1 and lmeTerminationDate is null and lmeHomeProductionDepartmentID = 'PROD' "
    


    
    df2 = pd.read_sql(notClockedOn, conn)
    
    df4 = pd.read_csv('testLeave.csv')
    df5 = pd.read_sql(leavesToday, conn)

    df5 = df5.drop_duplicates(subset=["lmeEmployeeName"], keep='first') #In case of duplicate employees showing up

    
    
    df2.to_csv('./quickTesting/df2NotClockedOutput.csv', index=True)
    df5.to_csv('./quickTesting/dfNotClockedOutput.csv', index=True)

    df2 = pd.concat([df2.set_index('Not_Clocked_Employee_Name'),df5.set_index('lmeEmployeeName')], axis=1, join='outer') #Joins if existing reason



    

    

   
    df2.index.names = ['Employee'] #Renames the index, you cannot use result.rename sadly
    df2.reset_index().rename(columns={'index': ''})

    df2 = df2.rename(columns={'lmpLeaveAccrualID':'Leave'})
    df2.replace({"LWOP":"Absent", "AL":"Absent", "ALA":"Absent", "WCOV":"Absent"}, inplace=True)


    df2.to_csv("testLeave.csv", index=True) #change index to true/false if you want it or if you want it gone

    df2 = pd.read_csv("testLeave.csv")
    
    
    return df2.to_dict('records2')





app.run_server(debug=True)

