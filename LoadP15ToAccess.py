import pyodbc
class P15File():
    """
    Structure is [{fldName:val, fldName:Val,..........}]
    Note removal of leading and trailing blanks on all values
    """
    def __init__(self, p15DataFN, p15HeaderFN):
        p15DataFile = open(p15DataFN,'r')
        p15HeaderFile = open(p15HeaderFN,'r')
        i=0
        self.header = []
        for rec in p15HeaderFile:
            if i == 0:
                self.header = rec.split('\t')
                i+=1
            else: break
        i=0
        self.data = []
        for rec in p15DataFile:
            i+=1
            line = rec.split('\t')
            lineDict = {}
            for fld in self.header:
                fldPos = self.header.index(fld)
                lineDict[fld] = line[fldPos].strip()
            self.data.append(lineDict)
        print('Lines imported = '+repr(i))
        p15DataFile.close()
        p15HeaderFile.close()

    def applyRecordFilters(self):
        recNo, delRecs, delRecs1 = 0, 0, 0
        copy =[]
        for rec in self.data:
            if rec['posting date'][0:4] not in ['2012','2013']:
                delRecs +=1
            elif rec['distribution channel'].find('Reseller') == -1:
                delRecs1 +=1
            else:
                copy.append(rec)
            recNo += 1
        print('deleted '+repr(delRecs) +' for invalid dates.')
        print('deleted '+repr(delRecs1) +' for invalid distribution channel.')
        self.data = copy

    def substitute(self,xlFile):
        conn=pyodbc.connect('Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)}; Dbq='+xlFile, autocommit=True)
        cursor = conn.cursor()
        #Create table from Excel in Access
        xlsStore = {}
        for row in cursor.execute('SELECT material, swapMaterial FROM [Missing Materials$]'):
            xlsStore[row.material] = row.swapMaterial
        cursor.close()
        conn.close()
        i=0
        for rec in self.data:
            rec['salesorg'] = '0110'
            rec['company'] = '0110'
            rec['plant'] = '1674'
            if rec['material'].strip() in xlsStore.keys():
                rec['material'] = xlsStore[rec['material'].strip()]

    def substituteShTo(self,xlFile):
        conn=pyodbc.connect('Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)}; Dbq='+xlFile, autocommit=True)
        cursor = conn.cursor()
        #Create table from Excel in Access
        xlsStore = {}
        for row in cursor.execute('SELECT P15ShipTo, SAPNoProposal FROM [Sheet1$]'):
            xlsStore[row.P15ShipTo.strip().zfill(10)] = row.SAPNoProposal.strip().zfill(10)
        cursor.close()
        conn.close()
        i=0
        for rec in self.data:
            if rec['ship-to custno'].strip() in xlsStore.keys():
                rec['ship-to custno'] = xlsStore[rec['ship-to custno'].strip()]
                i += 1
        print("Ship To converted on "+repr(i)+" records.")

    def flatten(self):
        flattened = []
        for line in self.data:
            li = []
            for fld in self.header:
                li.append(line[fld])
            flattened.append('\t'.join(li))
        return flattened

    def toSAFormat(self):
        SA = []
        for line in self.data:
            li = []
            li.append(line['salesorg'])
            li.append(line['sold-to custno'])
            li.append(line['material'])
            li.append(line['invoice number'])
            li.append('') #ZBILL_TYPE
            dt = line['posting date']
            li.append(dt[6:8]+dt[4:6]+dt[0:4])
            li.append('') #ITEM_CATEG
            iQty = line['invoice qty']
            if iQty.endswith('-'): li.append('-'+iQty.rstrip('-'))
            else: li.append(iQty)
            li.append(line['sales um'])
            li.append('') #ZRPT_QTY
            li.append('USD') #LOC_CURRCY
            li.append(line['plant'])
            li.append('') #PAYER
            li.append(line['ship-to custno'])
            li.append('0.00') #ZBVALLOC
            sAm = line['sales amt']
            if sAm.endswith('-'): sAmConv = '-'+sAm.rstrip('-')
            else: sAmConv = sAm
            nAm = line['net amt']
            if nAm.endswith('-'): nAmConv = '-'+nAm.rstrip('-')
            else: nAmConv = nAm
            li.append(sAmConv) #ZAINDCLOC
            li.append(nAmConv) #ZGVAL_LOC
            li.append(nAmConv) #ZAREBLOC
            li.append(nAmConv) #ZPKTSLLOC
            li.append('Y') #ZSAPCUSTI
            SA.append(';'.join(li))
        return SA

########### Main Program #################
path = 'C:\\DATA\\USERS\\U104675\\My Data\\P15 to P20\ProductionRun\\'
#Text file preparation
p15File = path + 'p15_sales_13052013.txt'
p15HeaderFile = path + 'P15_Headers_withtabs.txt'
missingMaterials = path + 'Missing Materials.xlsx'
missingShipTos = path + 'Unknown P15 Ship to Customers CURRENT.xlsx'
outP15 = path + 'convertedP15-v1.txt'
outSA = path + 'convertedSA-v1.txt' # Note new version number for ship to change
outP = open(outP15,'w')
outS = open(outSA,'w')

#Main
P15 = P15File(p15File,p15HeaderFile) #Import initial data
P15.applyRecordFilters()
P15.substitute(missingMaterials)
P15.substituteShTo(missingShipTos)

#Output result
lines = P15.flatten()
for line in lines: outP.write(line+'\n')
lines = P15.toSAFormat()
for line in lines: outS.write(line+'\n')
outP.close()
outS.close()
