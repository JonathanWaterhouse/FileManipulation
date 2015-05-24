## Initialisations
fieldSep = '\t' #Incoming files from Excel (since easy to output tab separated)
keySep = '¬' #The key used to store data internally. A join of what we want finally
path = 'C:\\Documents and Settings\\u104675\Desktop\\Chapter 11\\WW Revenue EKBI to SEM-Avishay - 2012 Q2\\'        
files = ['Avishay.txt','EKBI.txt']
correctionsFile = 'corrections.txt' 
accountsFile = 'GlobalMainAccount.txt'
kfDict = {}
AvishayDict = {}
ekbiDict = {}
data = [0,0,0,0,0,0,0] # Initialised with one zero for each key figure column
kfColsToCheck = [0,1,2,3,4,5,6] #Spreadsheet columns to do compares on. 0 is first key figure column
keyCols = [0,1,2] #location in columns of the row key elements (need concatenation to be full key)
keyPos = 3 #Offset of key fig cols due to key cols eg key in first col, offset is 1
colNames = ["Adjusted Pocket Sales","Total Basic Cogs","Total Distribution",
            "Total SIC","Total Other Cogs","Total Service Cost","Gross Profit"]
##Aggregate repeated key figures and store
for fName in files:
    f = open(path+fName,'r')
    count = 0
    for line in f:
        if count == 0: pass #header with column names expected in each input file
        else:
            keyy = ''
            for col in keyCols: keyy = keyy + line.split(fieldSep)[col]+keySep
            key = keyy.rstrip(keySep)
            for col in kfColsToCheck: #Loop over spreadsheet columns
                try:
                    if line.split(fieldSep)[col+keyPos] == "": data[col] = 0
                    else : data[col] = float(line.split(fieldSep)[col+keyPos].strip('"').replace(',',''))
                except (ValueError):
                    print (fName+ " Errored field at line "+repr(count))
                    continue
                curValList = kfDict.get(key,[0,0,0,0,0,0,0])
                newVal = data[col] + curValList[col]
                curValList[col] = newVal
                kfDict[key] = curValList
                curValList = [0,0,0,0,0,0,0]      
        count += 1
        #if count > 1000: break
    print(fName +' File line count :'+repr(count))
    f.close()
    if fName == 'Avishay.txt' : AvishayDict = dict(kfDict)
    else : ekbiDict = dict(kfDict)
    kfDict = {}

## Print checksums
for col in kfColsToCheck: #Loop over spreadsheet columns
    sum0 = 0 
    for el in AvishayDict.keys():
        sum0 = sum0 + AvishayDict[el][col]
    print("Avishay File Total in col "+colNames[col]+': ' + '%.2f' %(sum0))
    sum1 = 0
    for el in ekbiDict.keys():
        sum1 = sum1 + ekbiDict[el][col]
    print("EKBI File Total in col "+colNames[col]+": " + '%.2f' %(sum1))
    print('Correction to EKBI (Avishay - EKBI) for '+colNames[col]+": "+'%.2f' %(sum0-sum1))
##Compare files
setAvishay= set(AvishayDict.keys())
setEKBI = set(ekbiDict.keys())
setKeys = setAvishay.union(setEKBI)
#Iterate over keys
corrections = {}
curValList = [0,0,0,0,0,0,0]
for key in setKeys:
    for col in kfColsToCheck: #Loop over fey figure columns
        ## No repeated keys at this stage
        if key in setAvishay and key in setEKBI: #Correct EKBI
            avKF = float(AvishayDict[key][col])
            ekbiKF = float(ekbiDict[key][col]) 
            diff = avKF - ekbiKF
            curValList = corrections.get(key,[0,0,0,0,0,0,0])
            curValList[col] = '%.2f' %(diff)
            corrections[key] = curValList
            curValList = [0,0,0,0,0,0,0]                             
        if key in setAvishay and key not in setEKBI: #Add record to EKBI  
            avKF = float(AvishayDict[key][col])
            curValList = corrections.get(key,[0,0,0,0,0,0,0])
            curValList[col] = '%.2f' %(avKF)
            corrections[key] = curValList
            curValList = [0,0,0,0,0,0,0]
        if key not in setAvishay and key in setEKBI: #Add -ve of record to EKBI
            ekbiKF = float(ekbiDict[key][col]) 
            curValList = corrections.get(key,[0,0,0,0,0,0,0])
            curValList[col] = '%.2f' %(-1 * ekbiKF) 
            corrections[key] = curValList
            curValList = [0,0,0,0,0,0,0]    
#Tidy up to save space
AvishayDict, EKBIDict = {}, {}
setAvishay.clear()
setEKBI.clear()
setKeys.clear()
## Set up for output
print ("Completed..........outputting results to file")
output = open(path+correctionsFile,'w')
out=[]
## Add extra data to the key (Global Main Account)
accounts = file(path+accountsFile,'r')
globAcct = {}
count = 0
for line in accounts:
    globAcct[line.split(fieldSep)[0]] = line.split(fieldSep)[3]

#for k in globAcct.keys(): out.append("Customer : "+k+" has main acct : "+globAcct[k])
## Output corrections results to file
for k in corrections.keys(): 
    #Ignore any lines where all kf's are zero
    zero = True
    for kf in corrections[k]:
        if kf != '0.00' and kf != '-0.00': 
            zero = False
            break
    if not zero:  
        firstHyphen = k.find(keySep) #Find customer in the key field
        secondHyphen = k.find(keySep,firstHyphen+1)
        customer = k[firstHyphen+2:secondHyphen]
        try:
            key = k + keySep + globAcct[customer]
        except (KeyError): 
            key = k + keySep + 'NoGlobalAcct'
            count += 1
        out.append(key + fieldSep +','.join(corrections[k]))
out.sort()
out.insert(0,','+','.join(colNames))
for item in out: output.write(item+'\n')
output.close()
