import pyodbc

def importAllMaster(xlFile,DBFile):
    conn=pyodbc.connect('Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)}; Dbq='+xlFile, autocommit=True)
    cursor = conn.cursor()    
    #Create table from Excel in Access
    for row in cursor.tables():
        print row.table_name
    
    xlsStore = []
    for row in cursor.execute('SELECT * FROM [All Master$]'):
        print (row.Description)
        xlsStore.append([row.CPU,row.SCHED,row.JOB,row.FOLLOWS,row.ON,row.AT,row.OPENS,  row.NEEDS,row.Jobs,row.Description,row.UNTIL])
       
    cursor.close()
    conn.close()
    
    #Access
    conn = pyodbc.connect('DRIVER={Microsoft Access Driver (*.mdb)};DBQ='+DBfile)
    cursor = conn.cursor()
    try:
        SQL = "DROP TABLE [All Master-Import]"
        cursor.execute(SQL)
        conn.commit()
    except: print('INFORMATION: All Master-Import table did not exist. As expected, attempt to delete it failed'+'\n')
    cursor.execute("""
    CREATE TABLE [All Master-Import] (id COUNTER CONSTRAINT C1 PRIMARY KEY, CPU MEMO, SCHED MEMO, JOB MEMO, FOLLOWS MEMO, ONN MEMO, AT MEMO, OPENS MEMO, NEEDS MEMO, Jobs MEMO, Description MEMO, UNTIL MEMO);
    """)
    conn.commit()
    #Create table from Excel in Access
    for row in xlsStore:
        cursor.execute("""
        INSERT INTO [All Master-Import] (CPU,SCHED,JOB,AT,FOLLOWS,ONN,Jobs,Description,OPENS,NEEDS,UNTIL) 
        VALUES (?,?,?,?,?,?,?,?,?,?,?)
        """, [row[0],row[1],row[2],row[5], row[3],row[4],row[8],row[9],row[6],row[7],row[10]])    
    conn.commit()
    cursor.close()
    conn.close()    

def importSEMSchedules(xlFile,DBFile):
    conn=pyodbc.connect('Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)}; Dbq='+xlFile, autocommit=True)
    cursor = conn.cursor()        
    xlsStore = []
    for row in cursor.execute('SELECT * FROM [Sheet1$]'):
        print (row.Description)
        xlsStore.append([row.Schedule,row.Description,row[2],row.FOLLOWS])
       
    cursor.close()
    conn.close()
    
    #Access
    conn = pyodbc.connect('DRIVER={Microsoft Access Driver (*.mdb)};DBQ='+DBfile)
    cursor = conn.cursor()
    try:
        SQL = "DROP TABLE [SEM Schedules-Import]"
        cursor.execute(SQL)
        conn.commit()
    except: print('INFORMATION: All Master-Import table did not exist. As expected, attempt to delete it failed'+'\n')
    cursor.execute("""
    CREATE TABLE [SEM Schedules-Import] (id COUNTER CONSTRAINT C1 PRIMARY KEY, Schedule MEMO, description MEMO, RunsOn MEMO, FOLLOWS MEMO);
    """)
    conn.commit()
    #Create table from Excel in Access
    for row in xlsStore:
        cursor.execute("""
        INSERT INTO [SEM Schedules-Import] (Schedule,Description, RunsOn,FOLLOWS) 
        VALUES (?,?,?,?)
        """, [row[0],row[1],row[2],row[3]])    
    conn.commit()
    cursor.close()
    conn.close()        
    
def getJobs(schedule):
    cursor.execute("""
        SELECT SCHED, JOB, Jobs, Description FROM [All Master-Import] 
        WHERE SCHED = ? AND JOB IS NOT NULL
        """ , schedule)
    results = []
    for row in cursor.fetchall():
        results.append('\t\t' +repr(row.JOB)+'\t'+repr(row.Jobs)+'\t'+repr(row.Description))
    return results

def getScheduleDetail(schedule):
    cursor.execute("""
        SELECT DISTINCT Description FROM SchedDetails 
        WHERE Schedule = ?
        """ , schedule)
    description = cursor.fetchone()
    return repr(description)

def createTblSEMSchedules():
    try:
        SQL = "DROP TABLE [SEMSchedules]"
        cursor.execute(SQL)
        conn.commit()
    except: log.write('INFORMATION: SEMSChedules table did not exist. As expected, attempt to delete it failed'+'\n')
    cursor.execute("""
    CREATE TABLE SEMSchedules (id COUNTER CONSTRAINT C1 PRIMARY KEY, Schedule CHAR, description CHAR, RunsOn CHAR, FOLLOWS CHAR);
    """)
    conn.commit()
    cursor.execute("""
    INSERT INTO SEMSchedules (Schedule,description,RunsOn,FOLLOWS)
    SELECT MID(Schedule,INSTR(Schedule,'#')+1,LEN(Schedule)-INSTR(Schedule,'#')),Description,RunsOn,FOLLOWS
    FROM [SEM Schedules-Import];
    """)
    conn.commit() 
    
def createTblSchedDetails(srcFile): 
    try:
        SQL = "DROP TABLE [SchedDetails]"
        cursor.execute(SQL)
        conn.commit()
    except: log.write('INFORMATION: SchedDetails table did not exist. As expected, attempt to delete it failed'+'\n')
    cursor.execute("""
    CREATE TABLE [SchedDetails] (id COUNTER CONSTRAINT C1 PRIMARY KEY, Schedule CHAR, Description CHAR, Timing CHAR, When CHAR);
    """)
    conn.commit()
    schedTexts = {}
    text, at, when = '', '', ''
    data1 = open(path+srcFile,'r')
    for line in data1:
        if ('#' in line[0:2]) and ('T' in line[0:6]) and ('==' in line):
            text = line
        if line[0:8] == 'SCHEDULE' :
            start = line.index('#')+1
            end = len(line)+1
            thisSched = line[start:end].strip() 
        if line[0:2] == 'AT': at = line[3:]
        if line[0:3].count('ON')>0: when = line[3:].replace('"','',-1)
        if line[0:3] == 'END':
            schedTexts[thisSched] = text.strip('\n')
            cursor.execute("""
            INSERT INTO SchedDetails (Schedule, Description, Timing, When) 
            VALUES (?,?,?,?)
            """, [thisSched,text.strip('\n'), at.strip('\n'), when.strip('\n')])
            text, at, when = '','',''
    cursor.commit()
    data1.close()  
    
def createTblSchedJobs():
    try:
        SQL = "DROP TABLE [Sched-Jobs]"
        cursor.execute(SQL)
        conn.commit()
    except: log.write('INFORMATION: Sched-Jobs table did not exist. As expected, attempt to delete it failed'+'\n')
    cursor.execute("""
    CREATE TABLE [Sched-Jobs] (id COUNTER CONSTRAINT C1 PRIMARY KEY, Schedule CHAR, Job CHAR);
    """)
    conn.commit()
    
    cursor.execute("""
    INSERT INTO [Sched-Jobs] (Schedule, Job)
    SELECT SCHED, JOB FROM [All Master-Import]
    WHERE JOB IS NOT NULL
    """)
    conn.commit()
    
def createTblSchedSched(srcFile):
    try:
        SQL = "DROP TABLE [Sched-Sched]"
        cursor.execute(SQL)
        conn.commit()
    except: log.write('INFORMATION: Sched-Sched table did not exist. As expected, attempt to delete it failed'+'\n')
    cursor.execute("""
    CREATE TABLE [Sched-Sched] (id COUNTER CONSTRAINT C1 PRIMARY KEY, Schedule CHAR, Follows CHAR);
    """)
    conn.commit()
    # -> Get dependencies and write to table.
    data1 = open(path+srcFile,'r')
    follows = []
    for line in data1:
        if line[0:8] == 'SCHEDULE' :
            start = line.index('#')+1
            end = len(line)+1
            thisSched = line[start:end].strip() 
        if line[0:7] == 'FOLLOWS': 
            f = line[line.index('#')+1:line.index('.')].strip('\n')
            follows.append(f)
        if line[0:3] == 'END':
            for el in follows:
                cursor.execute("""
                INSERT INTO [Sched-Sched] (Schedule, Follows) VALUES(?,?)
                """, [thisSched,el])            
            follows = []
    conn.commit() 
    data1.close()
########### Main Program #################
path = 'C:\\DATA\\USERS\\U104675\\My Data\\DivestitureSchedules\\'

#MS Access preparation
DBfile = path + 'Schedule.mdb'
conn = pyodbc.connect('DRIVER={Microsoft Access Driver (*.mdb)};DBQ='+DBfile)
cursor = conn.cursor()

#Text file preparation
logFile = 'log.txt'
log = open(path+logFile,'w')
wizzbang = 'Prod-GDWScheds-Feb 25.xlsm'
SEMScheduleFile = 'SEM Schedules.xlsx'

#Import All Master
importAllMaster(path+wizzbang,DBfile)
#Import SEM Schedules
importSEMSchedules(path+SEMScheduleFile,DBfile)

#Create new SEM schedules table with clean key
createTblSEMSchedules()
    
#Get all schedule details REQUIRES text file extract  of Wizzbang tab 'GDWScheds'
createTblSchedDetails('schdGDWScheds.txt')

#Create a schedule to schedule dependency table in the database
createTblSchedSched('schdGDWScheds.txt')

#Create a schedule to Job dependency table
createTblSchedJobs()

#Output dependency information
log.write('>>>>>>>Dependencies and jobs for SEM schedules<<<<<<<<<<<\n')
SEMSchedules = []
SQL = "SELECT Schedule FROM SEMSchedules ORDER BY Schedule;"
for sch in cursor.execute(SQL): SEMSchedules.append(sch)
for sch in SEMSchedules:
    jobs = getJobs(sch)
    log.write(repr(sch)+'\t'+getScheduleDetail(sch)+'\n')
    for j in jobs: log.write('...............CONTAINS : '+j+'\n')
    subSched = []
    cursor.execute("""
    SELECT DISTINCT Follows FROM [Sched-Sched] WHERE Schedule = ?
    """, sch)
    for fol in cursor: subSched.append(fol)
    for ss in subSched:
        jobs = getJobs(ss)
        log.write('.........FOLLOWS : '+ repr(ss)+'\t'+getScheduleDetail(ss)+'\n')
        for j in jobs: log.write('...............CONTAINS : '+j+'\n')

#Output list of SEM schedules
log.write('>>>>>>>Required Schedules for SEM<<<<<<<<<<<\n')
SQL = "SELECT S.Schedule, SS.Follows FROM SEMSchedules AS S LEFT OUTER JOIN [Sched-Sched] AS SS ON S.Schedule = SS.Schedule;"
s = set()
SEMAllScheds = {}
for row in cursor.execute(SQL): 
    s.add(row.Schedule)
    s.add(row.Follows)
for sch in s:
    cursor.execute("""
    SELECT Description FROM SchedDetails WHERE Schedule = ?
    """,sch)
    text = cursor.fetchone()
    SEMAllScheds[sch] = text
schKeys = SEMAllScheds.keys()
schKeys.sort()
for key in schKeys:
    log.write(repr(key)[0:15] + '\t' + repr(SEMAllScheds[key])+'\n')  
    
#Output List Of resources
log.write('>>>>>>>Required Resources for SEM<<<<<<<<<<<\n')
jobs=[]
for key in schKeys: # loaded above
    for el in getJobs(key): jobs.append(el)
for j in jobs: log.write(j+'\n')

#Clean up
cursor.close()
conn.close()
log.close()