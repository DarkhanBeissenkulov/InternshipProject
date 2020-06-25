import tkinter as tk, os, sys, pandas as pd, time, math, multiprocessing, numpy as np, threading, queue
from pandasql import sqldf
from tkinter.filedialog import askopenfilename, askdirectory
from tkinter.ttk import Progressbar as Progressbar, Combobox
from tkinter import messagebox
from datetime import timedelta

# Define variable:
# Check points and description
CP1_D = 'Sheet name = Exception 1 - description comment ...'
CP2_D = 'Sheet name = Exception 2 - description comment ...'
CP3_D = 'Sheet name = Exception 3 - description comment ...'
#pysqldf = lambda q: sqldf(q, globals())
cnt = 0
maxn = int
percentage = 0
global progress, window
#Define max core and minus 1, in order to avoid totally computer down
num_workers = multiprocessing.cpu_count()
DPath = IPath = os.getcwd()

def ChooseFile(x,z):
    global IPath
    txt = askopenfilename(initialdir = IPath, title=z, filetypes = [("CSV file",".csv")])
    if txt:
        x = x+'.config(text = txt)'
        exec(x)
        IPath = os.path.split(txt)[0]
        #print('The next file chosen ' + txt)

def ChooseFile1(x,z):
    global IPath
    txt = askopenfilename(initialdir = IPath, title=z, filetypes = [("XLS/XLSX file",".xls .xlsx")])
    if txt:
        x = x+'.config(text = txt)'
        exec(x)
        IPath = os.path.split(txt)[0]
        #print('The next file chosen ' + txt)

def ChooseFolder(x):
    global DPath, IPath
    txt = askdirectory(initialdir = IPath)
    if txt:
        x = x+'.config(text = txt)'
        exec(x)
        DPath = IPath = txt
        #print('The next folder chosen ' + DPath)

def ReadFile (x):
    try:
        z = pd.read_excel(x)
        window.update()
        return z
    except Exception as e:
        messagebox.showerror("ERROR", e)
        exit(0)

def multiprocessing_func_new(inp_list_DF, ret1, ret2):
    try:
        #t= threading.current_thread()
        # Adjasting dataframes
        DF_VF = inp_list_DF[0]
        DF_WN_INTERNET = inp_list_DF[1]
        DF_WN_VOICE = inp_list_DF[2]
        DF_WN_IP = inp_list_DF[3]
        DF_WN_VP = inp_list_DF[4]
        # This creates dataframe for final extractions
        #Dataframe for exception 1 look for above CP1_D comments
        df_ex1 = pd.DataFrame(columns = ['Column1', 'Column2', 'and etc'])
        # Dataframe for exception 2 look for above CP2_D comments
        df_ex2 = pd.DataFrame(columns = ['Column1', 'Column2', 'and etc'])

        cond_join= '''
            select cast(f.[Column1] as INT) as Column1, 
                strftime('%d/%m/%Y', date(substr(f.[Date],5,4)||'-'||substr(f.[Date],3,2)||'-'||substr(f.[Date],1,2))) as "Date", 
                f.[Column2] as "Column2", f.[Column4] as "Column4",
                            r.[Column5] as "Column5", r.[Column6] as "Column6", r.[Column7], r.[Column8] as "Column8",
                            rr.[Column9] as "Column9", rr.[Column10] as "Column10",
                            wnp.[Column11] as "Column11",
                            wv.[Column12] as "Column12",
                            case 
                                when 
                                    ((f.[Column2] not like '%Rule%' 
                                                    and f.[Column2] not like '%Rule%' 
                                                    and f.[Column2] not like '%Rule%' 
                                                    and f.[Column2] not like '%Rule%'
                                                    ) and
                            not exists (
                                select 1
                                from DF_VF f1
                                join DF_WN_INTERNET r1 on f1.[Column1] = SUBSTR(r1.[Column],1, INSTR(r1.[Column], '@')-1)
                                join DF_WN_IP r2      on r2.[Column] = r1.[Column]
                                where f1.[Column1] = f.[Column1]
                                and (
                                    (r1.[Column] = 'active'
                                    and r2.[Column] = f1.[Column2]
                                    and f1.[Column4] > 0) 
                                    or
                                    (r1.[Column] = 'active'
                                        and r1.[Plan Name] = 'Rule'
                                        and f1.[Column2] = 'Rule'
                                    )
                                    or
                                    (f1.[Column2] = r2.[Rule]
                                    and exists (
                                        select 1 
                                          from DF_VF f2
                                         where f1.[Column1] = f2.[Column1]
                                           and f2.[Column2] = 'Rule'
                                        )
                                    and exists (
                                        select 1
                                          from DF_WN_VOICE v2
                                         where f1.[Column1] = v2.[Column]
                                           and v2.[Column] = 102
                                           and v2.[Column] = 'active'
                                        )
                                    )
                                )
                                and (f1.[Column2] not like '%Rule%' 
                                                    and f1.[Column2] not like '%Rule%' 
                                                    and f1.[Column2] not like '%Rule%' 
                                                    and f1.[Column2] not like '%Rule%'
                                                    )
                            )
                        )
                                    then 'not matches plan'
                                when 
                                    ( (f.[Column2] not like '%Rule%' 
                                                    and f.[Column2] not like '%Rule%' 
                                                    and f.[Column2] not like '%Rule%' 
                                                    and f.[Column2] not like '%Rule%'
                                                    ) and
                                not exists (
                                    select 1
                                    from DF_VF f1
                                    join DF_WN_INTERNET r1 on f1.[Column1] = SUBSTR(r1.[Column],1, INSTR(r1.[Column], '@')-1)
                                    join DF_WN_IP r2      on r2.[Column] = r1.[Column]
                                    where f1.[Column1] = f.[Column1]
                                    and r1.[Column] = 'active'
                                    and f1.[Column2] = case 
                                                                when length(r2.[Column]) > 1 then r2.[Column]
                                                                when length(r2.[Column]) <= 1 then f1.[Column2]
                                                            end
                                    and f1.[Column4] > 0
                                    and (f1.[Column2] not like '%Rule%' 
                                                    and f1.[Column2] not like 'Rule' 
                                                    and f1.[Column2] not like 'Rule' 
                                                    and f1.[Column2] not like 'Rule'
                                                    )
                                )
                        )
                                    then 'not matches ...'
                                when 
                                    ( f.[Column2] like '%Rule%' 
                    and f.[Column2] not like 'Rule'
                    and f.[Column2] not like 'Rule'
                    and 
                            not exists (
                                select 1 
                                from DF_WN_Rule    as wv
                                join DF_WN_VP       as wnp on wv.[Column] = wnp.[Column]
                                where f.[Column1] = wv.[Column]
                                and f.[Column2] = wnp.[Column]
                                and wv.[Column] = 'active'
                            )
                    
                )
                                    then 'not matching ...'
                                when 
                                    (length(cast(f.[Date] AS INT)) == 8 
                        and cast(substr(f.[Date],1,2) AS INT) >= 15 
                        and f.[Column2] not like 'Rule' 
                        and (f.[Column2] like 'Rule' or f.[Column2] like 'Rule')
                    )  
                                    then 'not macthces Part month'
                                else null
                            end as "Reason"
            from DF_VF          as F 
        left join DF_WN_INTERNET as r   ON f.[Column1] = SUBSTR(r.[Column],1, INSTR(r.[Column], '@')-1)
        left join DF_WN_IP       as rr  on rr.[Column] = r.[Column]
        left join DF_WN_VOICE    as wv  on f.[Column1] = wv.[Column]
                                            and wv.[Column] = 'active'
        left join DF_WN_VP       as wnp on wv.[Column] = wnp.[Column]
            where
            f.[Column2] not null
            and f.[Column] is null
            and r.[Column] = 'active'
            and (--not matches plan
                    ((f.[Column2] not like '%Rule%' 
                                                    and f.[Column2] not like 'Rule' 
                                                    and f.[Column2] not like 'Rule' 
                                                    and f.[Column2] not like 'Rule'
                                                    ) and
                            not exists (
                                select 1
                                from DF_VF f1
                                join DF_WN_INTERNET r1 on f1.[Column1] = SUBSTR(r1.[Column],1, INSTR(r1.[Column], '@')-1)
                                join DF_WN_IP r2      on r2.[Column] = r1.[Column]
                                where f1.[Column1] = f.[Column1]
                                and (
                                    (r1.[Column] = 'active'
                                    and r2.[Column] = f1.[Column2]
                                    and f1.[Column4] > 0) 
                                    or
                                    (r1.[Column] = 'active'
                                        and r1.[Plan Name] = 'RUle'
                                        and f1.[Column2] = 'Rule'
                                    )
                                    or
                                    (f1.[Column2] = r2.[Column]
                                    and exists (
                                        select 1 
                                          from DF_VF f2
                                         where f1.[Column1] = f2.[Column1]
                                           and f2.[Column2] = 'Rule'
                                        )
                                    and exists (
                                        select 1
                                          from DF_WN_VOICE v2
                                         where f1.[Column1] = v2.[Column]
                                           and v2.[Column] = 102
                                           and v2.[Column] = 'active'
                                        )
                                    )
                                )
                                and (f1.[Column2] not like 'Rule' 
                                                    and f1.[Column2] not like 'Rule' 
                                                    and f1.[Column2] not like 'Rule' 
                                                    and f1.[Column2] not like 'Rule'
                                                    )
                            )
                        )
                
                --not matches ...
                or ( (f.[Column2] not like 'Rule' 
                                                    and f.[Column2] not like 'Rule' 
                                                    and f.[Column2] not like 'Rule' 
                                                    and f.[Column2] not like 'Rule'
                                                    ) and
                                not exists (
                                    select 1
                                    from DF_VF f1
                                    join DF_WN_INTERNET r1 on f1.[Column1] = SUBSTR(r1.[Column],1, INSTR(r1.[Column], '@')-1)
                                    join DF_WN_IP r2      on r2.[Column] = r1.[Column]
                                    where f1.[Column1] = f.[Column1]
                                    and r1.[Column] = 'active'
                                    and f1.[Column2] = case 
                                                                when length(r2.[Column]) > 1 then r2.[Column]
                                                                when length(r2.[Column]) <= 1 then f1.[Column2]
                                                            end
                                    and f1.[Column4] > 0
                                    and (f1.[Column2] not like 'Rule' 
                                                    and f1.[Column2] not like 'Rule' 
                                                    and f1.[Column2] not like 'Rule' 
                                                    and f1.[Column2] not like 'Rule'
                                                    )
                                )
                        )
                
                --not matching ...
                or 
                ( f.[Column2] like 'Rule' 
                    and f.[Column2] not like 'Rule'
                    and f.[Column2] not like 'Rule'
                    and 
                            not exists (
                                select 1 
                                from DF_WN_VOICE    as wv
                                join DF_WN_VP       as wnp on wv.[Column] = wnp.[Column]
                                where f.[Column1] = wv.[Column]
                                and f.[Column2] = wnp.[Column]
                                and wv.[Column] = 'active'
                            )
                    
                )
                    
                --not macthces ...
                or (length(cast(f.[Date] AS INT)) == 8 
                        and cast(substr(f.[Date],1,2) AS INT) >= 15 
                        and f.[Column2] not like 'Rule' 
                        and (f.[Column2] like 'Rule' or f.[Column2] like 'Rule')
                    ) 
                )
        group by cast(f.[Column1] as INT), f.[Date], f.[Column2],
                r.[Column], rr.[Column], rr.[Column]
        order by f.[Column1], f.[Column4] desc;
            '''
        df_ex1 = sqldf(cond_join, locals())

        # Query when MSISDN not exists in WN report or that not active in WN
        cond_join= '''
            SELECT cast(f.[Column1] as INT) "Column1", 
            strftime('%d/%m/%Y', date(substr(f.[Date],5,4)||'-'||substr(f.[Date],3,2)||'-'||substr(f.[Date],1,2))) as "Date", 
            f.[Column2], f.[Column4],
            SUBSTR(n.[Column],1, INSTR(n.[Column], '@')-1) "Column", n.[Column5] as "Column5", n.[Column] as "Column", n.[Column] as "Column",
            r.[Column] as "Column"
        FROM DF_VF f
        JOIN DF_WN_INTERNET n
        ON f.[Column1] = SUBSTR(n.[Column],1, INSTR(n.[Column], '@')-1) 
        left join DF_WN_VOICE r on f.[Column1] = r.[Column]
        WHERE 
        
        f.[Column1] IS NULL or 
        (
        cast(f.[Column4] AS REAL) > 0
        and n.[Column] != 'active'
        and f.[Column2] not null
        and f.[Column] is null
        and 
            (
                (f.[Column2] not like 'Rule' 
                and f.[Column2] not like 'Rule'
                and (f.[Column2] like 'Rule' or f.[Column2] like 'Rule') 
                and not exists (select 1 
                                from DF_WN_INTERNET nn 
                            where SUBSTR(n.[Column],1, INSTR(n.[Column], '@')-1) 
                                = SUBSTR(nn.[Column],1, INSTR(nn.[Column], '@')-1)
                                and nn.[Column] = 'active'
                            )
                )
                        
            or (f.[Column2] like 'Rule'
                and f.[Column2] not like 'Rule'
                and not exists (select 1 
                                    from DF_WN_VOICE rr
                                where r.[Column] = rr.[Column]
                                    and rr.[Column] = 'active'
                                )
                )
            )
            and not exists (
                select 1
                    from DF_VF f1
                    where f1.[Column1] = f.[Column1]
                    and f1.[Column2] = f.[Column2]
                    and cast(f1.[Column4] AS REAL) < 0
                    and date(substr(f1.[Date],5,4)||'-'||substr(f1.[Date],3,2)||'-'||substr(f1.[Date],1,2)) >
                        date(substr(f.[Date],5,4)||'-'||substr(f.[Date],3,2)||'-'||substr(f.[Date],1,2))
            )
        )
        order by f.[Column1], f.[Column4] desc;
        '''
        df_ex2 = sqldf(cond_join, locals())
        #print(df_ex2.shape[0])

        # #Default is to use xlwt for xls, openpyxl for xlsx.
        #writer = pd.ExcelWriter(DPath +'sample1_'+t.getName()+'.xlsx', engine='openpyxl')
        # #workbook = writer.book
        # # Write each dataframe to a different worksheet.
        #df_ex1.to_excel(writer, sheet_name='Exception 1', index=False)
        #df_ok.to_excel(writer, sheet_name='OK', index=False)
        # result = 'return from '+t.getName()
        ret1.put(df_ex1)
        ret2.put(df_ex2)
        # Close the Pandas Excel writer and output the Excel file.
        # try:
        #     writer.save()
        # except Exception as e:
        #     messagebox.showerror("ERROR", e)
        #     exit(0)
    except Exception as e:
        messagebox.showerror("ERROR", e)
        exit(0)
    
    
    
    

def RunButton():
    global DPath, maxn, percentage, num_workers
    openFile.config(state='disabled')
    openFile2.config(state='disabled')
    openFile6.config(state='disabled')
    openFile3.config(state='disabled')
    openFile4.config(state='disabled')
    openFile5.config(state='disabled')
    RunB.config(state='disabled')
    #selections.config(state= 'disabled')
    
    # Read files
    try:
        DF_VF = pd.read_csv(lbl2.cget("text"))
        DF_VF.drop(DF_VF.columns.difference(['columns']), 1, inplace=True)
    except Exception as e:
        messagebox.showerror("ERROR", e)
        return    
    DF_WN_INTERNET = ReadFile(lbl.cget("text")) #pd.read_excel(lbl.cget("text"))
    # need to remove columns that are not used in the query due to an error
    # OverflowError: Python int too large to convert to SQLite INTEGER    
    try:
        DF_WN_INTERNET.drop(DF_WN_INTERNET.columns.difference(['Columns name']), 1, inplace=True)
    except Exception as e:
        messagebox.showerror("ERROR", e +'\nProbably, you chosen wrong file.')
        return
    DF_WN_VOICE = ReadFile(lbl5.cget("text"))
    try:
        DF_WN_VOICE.drop(DF_WN_VOICE.columns.difference(['Columns name']), 1, inplace=True)
    except Exception as e:
        messagebox.showerror("ERROR", 'Probably, you chosen wrong file.')
        return
    DF_WN_IP = ReadFile(lbl3.cget("text"))
    DF_WN_VP = ReadFile(lbl4.cget("text"))

    list_DF = [DF_VF, DF_WN_INTERNET, DF_WN_VOICE, DF_WN_IP, DF_WN_VP]

    window.update()

    # This creates dataframe for final extractions
    #Dataframe for exception 1 look for above CP1_D comments
    df_ex1 = pd.DataFrame(columns = ['Columns...'])
    # Dataframe for exception 2 look for above CP2_D comments
    df_ex2 = pd.DataFrame(columns = ['Columns...'])
    # Dataframe for exception 3 look for above CP3_D comments
    df_ex3 = pd.DataFrame(columns = ['Columns...'])

    # Defining dataframe for Unique Mobile Numbers
    dfmn = DF_VF.filter(['Column1'], axis=1).drop_duplicates()
    dfmn = dfmn.dropna().reset_index(drop=True)
    dfmn['Column1'] =  pd.to_numeric(dfmn['Column1'],downcast='integer',errors='coerce')
    dfmn = dfmn.reset_index(drop=True)

    # Defin how many MSISDN going to process
    maxn = dfmn.shape[0]
    

    window.update()

    # Query for the main sheet
    cond_join= '''
        SELECT f.[column], f.[column], f.[column], f.[column], f.[Column2], f.[Column4]
        FROM DF_VF f
        WHERE f.[Column2] in ('some rule')
    '''
    df_main = sqldf(cond_join, locals())
    #Add comments under the main sheet dataframe
    list = ['',CP1_D, CP2_D, CP3_D]
    for i in list:
        df_main = df_main.append({'Invoice Number':i}, ignore_index=True)

    # Query for searching On/Off TXT messages
    cond_join= '''
        SELECT f.[Column1], f.[column], f.[Column2], f.[column]
        FROM DF_VF f
        WHERE f.[Column2] in ('some rule')
    '''
    df_ex3 = sqldf(cond_join, locals())
    
    results1 = None
    #results2 = None
    ret1 = queue.Queue()
    ret2 = queue.Queue()
    
    #df_split = np.array_split(dfmn, num_workers)
    # for i in range(num_workers):
    #     x = threading.Thread(target=multiprocessing_func, args=(df_split[i:i+1], list_DF, ret1, ret2))
    #     x.start()
    #     time.sleep(0.1)
    start = time.time()
    x = threading.Thread(target=multiprocessing_func_new, args=(list_DF, ret1, ret2))
    x.start()
    
    
    th_list = threading.enumerate()
    main_thread = threading.current_thread()
    pb_value = 0
    direction = 1
    while threading.active_count() > 1:
        # THIS UPDATING TEXT IN PROGRESSBAR WITH PERCENTAGE        
        #percentage = round(cnt/maxn*100)  # Calculate percentage.
        #style.configure('SPB', text='Processing MSISDN:#' + str(cnt) + ' from ' + str(maxn) + '; percentage: ' + str(percentage) + '%')
        style.configure('SPB', text='During processing, please wait for completion')
        #progress['value'] = percentage
        if pb_value < 100 and direction == 1: pb_value += 20
        elif pb_value == 0 and direction == 2:
            pb_value +=20
            direction = 1
        elif pb_value == 100 and direction == 1:
            pb_value -= 20
            direction = 2
        elif pb_value > 0 and direction == 2: pb_value -= 20
        progress['value'] = pb_value
        window.update()
        time.sleep(0.5)
        # if t is main_thread:
        #     print (t.getName() + ' is still alive = ' + str(t.is_alive()))
    else:
        for t in th_list:
            if t is main_thread:
                continue
            results1 = ret1.get()
            df_ex2 = ret2.get()
            df_ex1 = pd.concat([df_ex1, results1], ignore_index=True)
            #df_ex2 = pd.concat([df_ex2, results2], ignore_index=True)
            #df_ok = pd.concat([df_ok, results2], ignore_index=True)            
            ret1.task_done()
            ret2.task_done()
            t.join()
    elapsed = str(timedelta(seconds=(time.time() - start))).split(".")[0]
    # style.configure('SPB', text='Processing MSISDN: #' + str(maxn) + ' from ' + str(maxn) + '; percentage: ' + str(100) + '%')
    # progress['value'] = 100
    # time.sleep(0.5)
    
        

    #https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.ExcelWriter.html
    #class pandas.ExcelWriter(path, engine=None, date_format=None, datetime_format=None, mode='w', **engine_kwargs)
    #Default is to use xlwt for xls, openpyxl for xlsx.
    FilePath = DPath +'\\text '+time.strftime("%H%M-%d%m%y")+'.xlsx'
    writer = pd.ExcelWriter(FilePath, engine='xlsxwriter')
    #workbook = writer.book
    # Write each dataframe to a different worksheet.
    df_main.to_excel(writer, sheet_name='Main', index=False)
    df_ex1.to_excel(writer, sheet_name='Exception 1', index=False)
    #df_ex2.to_excel(writer, sheet_name='Exception 2', index=False)
    #df_ex3.to_excel(writer, sheet_name='Exception 3', index=False)
    #df_ex4.to_excel(writer, sheet_name='Exception 4', index=False)
    df_ex2.to_excel(writer, sheet_name='Exception 2', index=False)
    df_ex3.to_excel(writer, sheet_name='Exception 3', index=False)
    #df_ok.to_excel(writer, sheet_name='OK', index=False)
    
    #Nice to have:
    #- adjust column width to be more user friendly (so that users don't have to do it manually every month)
    #- freeze the top row on every tab
    #- add filters to all Exception tabs
    worksheet = writer.sheets['Main']
    worksheet.freeze_panes(1, 0)
    worksheet.autofilter(0,0,worksheet.dim_rowmax, worksheet.dim_colmax)
    #worksheet.set_column(0, 0, 10.1)
    for i1, col1 in enumerate(df_main.iloc[:,0:]):
        column_len = df_main.iloc[:,i1].astype(str).str.len().max()
        if math.isnan(column_len):
            column_len = 0
        column_len = max(column_len, len(col1)) + 2
        worksheet.set_column(i1, i1, column_len)
    #measurer = np.vectorize(len)
    #list_df = ['df_ex1', 'df_ex2', 'df_ex3', 'df_ex4', 'df_ex5', 'df_ex6', 'df_ok']
    list_df = [df_ex1, df_ex2, df_ex3]
    list_sheets = ['Exception 1', 'Exception 2', 'Exception 3']
    #dictionary = dict(zip(list_df, list_sheets))
    #worksheet = writer.sheets['Exception 1']
    #worksheet.freeze_panes(1, 0)
    #worksheet.autofilter(0,0,worksheet.dim_rowmax, worksheet.dim_colmax)

    #for z, col in enumerate(df_ex1.columns):
    #    print(z, col)
    #    # find length of column i
    #    column_len = df_ex1.iloc[:,z].astype(str).str.len().max()
    #    # Setting the length if the column header is larger
    #    # than the max column value length
    #    column_len = max(column_len, len(col)) + 2
    #    # set the column length
    #    worksheet.set_column(z, z, column_len)
    
    for i2 in range(len(list_df)):
        list_cols = list_df[i2].columns
        worksheet = writer.sheets[list_sheets[i2]]
        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0,0,worksheet.dim_rowmax, worksheet.dim_colmax)
        for z, col in enumerate(list_cols):
            # find length of column i
            column_len = list_df[i2].iloc[:,z].astype(str).str.len().max()
            if math.isnan(column_len):
                column_len = 0
            # Setting the length if the column header is larger
            # than the max column value length
            column_len = max(column_len, len(col)) + 2
            # set the column length
            worksheet.set_column(z, z, column_len)
    #res1 = measurer(df_ex1.values.astype(str)).max(axis=0)
    #res2 = df_ex1.columns
    #for i1, col in enumerate(res1):
    #    print('i1, col')
    #    print(i1, col)
    #    #find length of column i len(res2[i])
    #    #Setting the length if the column header is larger
    #    #than the max column value length
    #    #set the column length
    #    worksheet.set_column(i1, i1, col+2)
    
 

    # Close the Pandas Excel writer and output the Excel file.
    try:
        writer.save()
    except Exception as e:
        messagebox.showerror("ERROR", e)
        exit(0)    
    
    # PopUp meassage of finishing
    messagebox.showinfo("Processing finished", "Processing successfully finished in "+ elapsed +".\nExcel file saved at path\n"+FilePath)
    #Clean up and update all labels 
    style.configure('SPB', text='Progress bar')
    progress['value'] = 0
    lbl2.config(text = lblVF)
    lbl.config(text = lblWN)
    lbl5.config(text = lblVS)
    lbl3.config(text = lblIR)
    lbl4.config(text = lblVR)
    lbl6.config(text = "Default destination folder is \n" + os.getcwd())
    openFile.config(state='active')
    openFile2.config(state='active')
    openFile6.config(state='active')
    openFile3.config(state='active')
    openFile4.config(state='active')
    openFile5.config(state='active')
    RunB.config(state='active')    
    #selections.config(state= 'readonly')
    window.update()

if __name__ == '__main__':
    global progress, window
    window = tk.Tk()
    window.geometry('750x430')
    window.title("Name")

    lblVF = "Please choose invoice detail csv file"
    lbl2 = tk.Label(window, text=lblVF, bd=1, relief="solid", width=60, height=3, padx=10, anchor='w', wraplength=400, justify="left")
    lbl2.grid(row=1,column=0)
    lblWN = "Please choose Internet Services xls file"
    lbl = tk.Label(window, text=lblWN, bd=1, relief="solid", width=60, height=3, padx=10, anchor='w', wraplength=400, justify="left")
    lbl.grid(row=2,column=0)
    lblVS = "Please choose Voice Services xls file"
    lbl5 = tk.Label(window, text=lblVS, bd=1, relief="solid", width=60, height=3, padx=10, anchor='w', wraplength=400, justify="left")
    lbl5.grid(row=3,column=0)
    lblIR = "Please choose an Internet Rules file"
    lbl3 = tk.Label(window, text=lblIR, bd=1, relief="solid", width=60, height=3, padx=10, anchor='w', wraplength=400, justify="left")
    lbl3.grid(row=4,column=0)
    lblVR = "Please choose a Voice Rules file"
    lbl4 = tk.Label(window, text=lblVR, bd=1, relief="solid", width=60, height=3, padx=10, anchor='w', wraplength=400, justify="left")
    lbl4.grid(row=5,column=0)
    lbl6 = tk.Label(window, text="Default destination folder is " + os.getcwd(), bd=1, relief="solid", width=60, height=3, padx=10, anchor='w', wraplength=400, justify="left")
    lbl6.grid(row=6,column=0)

    openFile = tk.Button(window, text="Choose File",  width=20, command= lambda: ChooseFile('lbl2',lblVF))
    openFile.grid(row=1,column=1, padx=10)

    openFile2 = tk.Button(window, text="Choose File",  width=20, command= lambda: ChooseFile1('lbl',lblWN))
    openFile2.grid(row=2,column=1, padx=10)

    openFile6 = tk.Button(window, text="Choose File",  width=20, command= lambda: ChooseFile1('lbl5',lblVS))
    openFile6.grid(row=3,column=1, padx=10)

    openFile3 = tk.Button(window, text="Choose File",  width=20, command= lambda: ChooseFile1('lbl3',lblIR))
    openFile3.grid(row=4,column=1, padx=10)

    openFile4 = tk.Button(window, text="Choose File",  width=20, command= lambda: ChooseFile1('lbl4',lblVR))
    openFile4.grid(row=5,column=1, padx=10)

    openFile5 = tk.Button(window, text="Choose Folder",  width=20, command= lambda: ChooseFolder('lbl6'))
    openFile5.grid(row=6,column=1, padx=10)

    RunB = tk.Button(window, text="Push to Run",  width=20, command= lambda: RunButton())
    RunB.grid(row=7,column=1, padx=10)

    sellabel = tk.Label(window, text='Please choose, corresponding files and output path', width=60, height=3, padx=2, anchor='w', wraplength=450, justify="left")
    sellabel.grid(row=0,column=0)

    # Progress bar widget 
    style = tk.ttk.Style(window)
    style.layout('SPB',
                [('Horizontal.Progressbar.trough',
                {'children': [('Horizontal.Progressbar.pbar',
                                {'side': 'left', 'sticky': 'ns'})],
                    'sticky': 'nswe'}),
                ('Horizontal.Progressbar.label', {'sticky': ''})])

    #progress = Progressbar(window, style='SPB', orient = 'horizontal', length = 600, mode = 'determinate') 
    progress = Progressbar(window, style='SPB', orient = 'horizontal', length = 600, mode = 'indeterminate') 
    progress.grid(row=9,columnspan=3, pady =10)
    style.configure('SPB', text='Progress bar')
    #selections.bind("<<ComboboxSelected>>", callbackFunc)

    tk.mainloop()


#print('Application stopped')
#input("Press enter to close the console window...")