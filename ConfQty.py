import sqlite3
__author__ = 'U104675'
class CRMConfQty():
    def getRecords(self,fileIn):
        """
        set up a data structure based on file input structure
        self._cols =  [colName, colName, .......]
        self._data = [value, value, ..........]
                ......}}
        """
        self._data = []
        for fi in fileIn:
            f = open(fi,'r')
            count = 0
            cols, lineSplit = [], []
            for line in f:
                # get field names
                if count == 0:
                    self._cols = [el.strip() for el in line.lstrip('| |').rstrip('|\n').split('|')]
                    lineLen = len(self._cols) #This defines how many columns we should have in the data
                # get data
                if count >= 1:
                    lineSplit = line.lstrip('|').rstrip('|\n').split('|')
                    #Sort out some peculiarities of initial bytes of each record
                    #pipeCnt, charCnt = 0, 0
                    #for char in line:
                    #    if char == '|': pipeCnt += 1
                    #    if char != '|' and char != '"' and not char.isspace():
                    #        line = line[charCnt:]
                    #        break
                    #    charCnt += 1
                    #lineSplit = line.rstrip('|\n').split('|')
                    #if pipeCnt >= 2:
                    #    for k in range(pipeCnt-2): lineSplit.insert(0,'')
                    #ln = len(lineSplit)
                    #Add missing columns due to dropped empty fields
                    #if ln < lineLen:
                    #    for pipeCnt in range(ln,lineLen): lineSplit.append('')
                    self._data.append([el.strip() for el in lineSplit])
                count += 1
                #if count > 100: break
            f.close()

    def uniqueValues(self):
        coorder = set()
        for line in self._data: coorder.add(line[self._cols.index('COORDER')])
        print(repr(len(coorder)) + ' unique vales of COORDER')

    def createPCATable(self,database):
        """
        Load data into a sqlite3 table
        """
        conn = sqlite3.connect(database)
        c = conn.cursor()
        #Delete table if already exists
        c.execute('''DROP TABLE IF EXISTS ZO00060''')
        #Create table
        c.execute('''CREATE TABLE ZO00060 (CURTYPE text, FUNC_AREA text, GL_ACCOUNT text, PCA_DOCNO text, PCA_ITEMNO text,
            COORDER text, CUSTOMER text, REFER_DOC text, REFER_ITM text, ZMAINACCT text, AC_DOC_TYP text, UNIT text,
            COMP_CODE text, QUANTITY real, SALES real )''')
        #Load data
        dataList = []
        for line in self._data:
            dataList.append(line[self._cols.index('CURTYPE')])
            dataList.append(line[self._cols.index('FUNC_AREA')])
            dataList.append(line[self._cols.index('GL_ACCOUNT')])
            dataList.append(line[self._cols.index('PCA_DOCNO')])
            dataList.append(line[self._cols.index('PCA_ITEMNO')])
            dataList.append(line[self._cols.index('COORDER')])
            dataList.append(line[self._cols.index('CUSTOMER')])
            dataList.append(line[self._cols.index('REFER_DOC')])
            dataList.append(line[self._cols.index('REFER_ITM')])
            dataList.append(line[self._cols.index('/BIC/ZMAINACCT')])
            dataList.append(line[self._cols.index('AC_DOC_TYP')])
            dataList.append(line[self._cols.index('UNIT')])
            dataList.append(line[self._cols.index('COMP_CODE')])
            dataList.append(line[self._cols.index('QUANTITY')])
            dataList.append(line[self._cols.index('SALES')])
            tup = tuple(dataList)
            dataList = []
            c.execute('INSERT INTO ZO00060 VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)', tup)
        conn.commit()
        conn.close()

    def countDistinct(self):
        conn = sqlite3.connect('myDb.db')
        c = conn.cursor()
        c.execute('''SELECT COUNT(DISTINCT COORDER) FROM ZO00060''')
        data = c.fetchone()
        conn.close()
        return data

    def calculate(self,db, fileout):
        """
        Analyse input data (which is a subset of rows and columns from ZO00060).
        The aim is to understand how well the data would support the assignment of CRM Confirmation Qty to
        Basic margin GL accounts given they appear in ZO00060 in other accounts. Essentially we:
        a) Select lines coming from CRM (COORDER non blank and starting "4"
        b) Find all lines with non zero qty (reporting any found where 0FUNC_AREA != "999"
        c) Find lines with GL_ACCOUNT starting "003" with the same dollar amounts as the lines with
        FUNC_AREA = "999" and non zero qtys and assign the qty values to the matched "0003" series lines
        Various error conditions are reported when it would not be possible to make a rational decision as to
        which "0003" series line to assign the qty to.
        """
        conn = sqlite3.connect(db)
        c = conn.cursor()
        # How many COORDERS
        c.execute('''SELECT COUNT(DISTINCT COORDER) FROM ZO00060''')
        data = c.fetchone()
        messages = []
        messages.append(' MSG01: No of COORDERS Examined : ' + repr(data))
        c.execute('''SELECT COORDER, FUNC_AREA, GL_ACCOUNT, QUANTITY, SALES FROM ZO00060
        WHERE substr(COORDER,1,1)='4'
        ORDER BY COORDER, substr(FUNC_AREA,1,3) desc, GL_ACCOUNT''')
        #LIMIT 1000''')
        rows = c.fetchall()
        fout = open(fileout,'w')
        curr_Coorder = ''
        coorder_Grp, src, tgt, problem = [],[],[],[]
        PCADocs, CoordersNo3Accts, unmatchedCoorderLines, countProbCO, cntQtyNotOn999  = 0, 0, 0, 0, 0
        qtyErrCnt = {}
        problemCO = False
        GL_accnts = set(['0003180100','0003180103','0003160100','0003160103','0003170100','0003170103'])
        GL_accnts_999 = set(['S5033','S5034','S5035','S5036','S5037','S5038','S5039','S5040','S5041','0006230600'])
        for row in rows:
            PCADocs += 1
            #TODO First COORDER is ignored
            if row[0] != curr_Coorder:
                ##Do some stuff with the lines we got for the last COORDER
                #Only consider COORDERs where there are postings to 'GCSS COGS accounts'
                if len([line[2] for line in coorder_Grp if line[2] in GL_accnts]) == 0:
                    messages.append(curr_Coorder + ' MSG10: WARNING: was ignored as had no GCSS COGS GL Accounts')
                    curr_Coorder = row[0]
                    coorder_Grp.append(row)
                    CoordersNo3Accts += 1
                    continue
                ##Store relevant information from COORDERS with GCSS COGS GL_ACCOUNTS
                countGCSS = 0
                for line in coorder_Grp:
                    #Store 3 series GL Accounts lines as target of quantities
                    if line[2] in GL_accnts:
                        tgt.append([line[2], line[4], '']) #amnt[ ['GL_ACCOUNT', USD value, qty],  etc...]
                        messages.append(curr_Coorder + ' MSG15-' + repr(countGCSS) + ': has GCSS Cost Line ' + repr(line))
                        countGCSS += 1
                    #Store GL Accounts where qty is non zero as source of quantities to be posted to target
                    if line[3] != 0.0 and line[2] in GL_accnts_999:
                        #qtyCount += 1
                        src.append([line[2], line[4], line[3]]) #amnt[ ['GL_ACCOUNT', USD value, qty],  etc...]
                        messages.append(curr_Coorder + ' MSG20: has line with a quantity ' + repr(src[-1]))
                        if len(src) == 0: messages.append(curr_Coorder + ' MSG21: ERROR: Qty on FUNC_AREA = '+ line[1])
                        if line[1] != '999':
                            messages.append(curr_Coorder + ' MSG30: ERROR: Qty on FUNC_AREA = '+ line[1])
                            cntQtyNotOn999 += 1
                #Loop over unique values of USD Value
                for val in set([line[1] for line in tgt]):
                    t = len([line[1] for line in tgt if line[1] == val]) #No. time val in target 3 series lines
                    s = len([line[1] for line in src if line[1] == val]) #No. time val in src lines (those with qtys)
                    if t > s:
                        messages.append(curr_Coorder + ' MSG40: ERROR: Value ' + repr(val) + ' appears with qty in source ' + repr(s)
                            + ' times but target has it only ' + repr(t) + ' times' )
                        #Count the number of each f these types of errors
                        try: qtyErrCnt[(s,t)] +=1
                        except (KeyError): qtyErrCnt[(s,t)] = 1
                        unmatchedCoorderLines  += 1
                        problemCO = True
                    elif t == s:
                        if len(set(line[0] for line in tgt if line[1] == val)) == 1:
                            # All GL Accounts for value in the target are the same we can post each value wherever we want
                            qtys = [line[2] for line in src if line[1] == val] # The qtys we need to distribute in the target
                            count, posted = 0, 0
                            for qty in qtys:
                                for line in tgt:
                                    if posted < s:
                                        if line[1] == val:
                                            line[2] = qty
                                            tgt[count] = line
                                            posted += 1
                                            messages.append(curr_Coorder + ' MSG50: Qty ' + repr(qty) + ' posted to ' + repr(line))
                                    count += 1
                        else:
                            messages.append(curr_Coorder + ' MSG60: ERROR: Value ' + repr(val) + ' appears with qty in source ' + repr(s)
                            + ' times and target has it ' + repr(t) + ' times, however target GL_Accounts are different so assignment decision was not made.' )
                            problemCO = True
                            unmatchedCoorderLines  += 1
                    else: #t<s
                        messages.append(curr_Coorder + ' MSG70: ERROR: target has ' + repr(t) + ' records but source has less '
                        + repr(s) + ' for value ' + repr(val))
                        problemCO = True
                        unmatchedCoorderLines  += 1
                messages.append(curr_Coorder + ' MSG99: END. \n')
                #Clean up
                coorder_Grp = []
                src, tgt = [], []
                if problemCO:
                    countProbCO += 1
                    problem.append(' MSG09: ' + curr_Coorder + ' has problems. Qty found but cannot match to GCSS COGS GL Account line.')
                problemCO = False
                #Accumulate first of next group of records
                curr_Coorder = row[0]
                coorder_Grp.append(row)
            else:
                coorder_Grp.append(row)
        conn.close()
        messages.append(' MSG00: Total PCA Documents from CRM transactions examined = ' + repr(PCADocs))
        messages.append(' MSG02: Total COORDERS with no GCSS COGS GL accounts in ZO00060 extract = ' + repr(CoordersNo3Accts))
        messages.append(' MSG03: Total COORDERS where we have a GCSS Cost GL Account line but cannot match to a quantity = ' + repr(countProbCO))
        messages.append(' MSG04: Total unmatched COORDER Line values = ' + repr(unmatchedCoorderLines))
        messages.append(' MSG05: Total COORDER lines where qty is on line with 0FUNC_AREA not "999" = ' + repr(cntQtyNotOn999))
        for k,v in qtyErrCnt.items():
            messages.append (' MSG06: Quantity Lookup error information in form {(source#,Target#):count} ' + repr(k) + ' : ' + repr(v))
        messages.extend(problem)
        messages.sort()
        for line in messages: fout.write(line + '\n')
        fout.close()

if __name__ == '__main__':
    folder = 'C:\Documents and Settings\\u104675\Desktop\Service Investigation\\'
    fileIn = []
    fileIn.append(folder + 'ZO00060_FEB_2013_0351_LOC.txt')
    fileIn.append(folder + 'ZO00060_FEB_2013_0351_USD.txt')
    fileIn.append(folder + 'ZO00060_FEB_2013_0286_LOC.txt')
    fileIn.append(folder + 'ZO00060_FEB_2013_0286_USD.txt')
    fileIn.append(folder + 'ZO00060_mar_ADJ.txt')
    fileOut = folder + 'ZO00060_RESULTS.txt'
    DbName = folder + 'ConfQtyAnalysis.db'
    c = CRMConfQty()
    c.getRecords(fileIn)
    #c.uniqueValues()
    c.createPCATable(DbName)
    #c.countDistinct()
    #c.calculate(DbName,fileOut)