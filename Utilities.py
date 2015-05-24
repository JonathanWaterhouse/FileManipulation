import sqlite3
class SAPFile:
    def fixLines1(self,filein,fileout):
        """
        This class takes a sap file from ZSE16 on a DSO in unconverted format. It
        reformats the data so that it can be loaded back via a flat file data source
        to the same DSO
        a) fixes quantities by removing trailing minus and adding to the front
        b) Puts dates in the correct format dd.mm.yyyy -> YYYYMMDD
        c) Fixes CALMONTH format YYYY.PPP -> PPPYYYY
        Field positions are referred to by number starting from zero and are hard
        coded
        """
        fin = open(filein,'r')
        fout = open(fileout,'w')
        l_cnt = 0
        for line in fin:
            l_List = line.rstrip('|').lstrip('|').split('|')
            l_List = [el.strip() for el in l_List]
            #Don't process header an further
            if l_cnt == 0:
                l_cnt += 1
                line_o = ",".join(l_List).rstrip(',') + '\n'
                fout.write(line_o)
                continue
            #Fix Quantities
            for index in [49, 51, 60]: #ZAMNT_GRP, ZAMNT_LOC, GLPCAQTY
                if l_List[index][-1] == '-': l_List[index] = '-' + l_List[index].rstrip('-')
                l_List[index] = l_List[index].replace(',','')
            #Fix dates
            for index in [0, 16]: #Posting date, calendar day
                temp = l_List[index][6:10] + l_List[index][3:5] + l_List[index][0:2]
                l_List[index] = "".join(temp)
            index = 17 #Cal year month
            tmp = l_List[index][4:6] + l_List[index][0:4]
            l_List[index] = "".join(tmp)
            index = 31 #Fiscal year period
            tmp = l_List[index][4:7] + l_List[index][0:4]
            l_List[index] = "".join(tmp)
            #Put it all back together again
            line_o = ",".join(l_List).rstrip(',') + '\n'
            fout.write(line_o)
        fout.close()

    def getData (self,fileIn):
        """
        Split lines by the normal ZE16 line separator "|" removing leading and trailing stuff
        """
        fin = open(fileIn,'r')
        l_out = []
        i = 0
        for line in fin:
            l_List = line.rstrip('\n').rstrip('|').lstrip('|').split('|')
            l_out.append(tuple(el.strip() for el in l_List))
        return tuple(el for el in l_out)

    def loadToSqlite(self, data):
        """
        Load data into a sqlite3 table
        """
        conn = sqlite3.connect('myDb.db')
        c = conn.cursor()
        #Delete table if already exists
        c.execute('''DROP TABLE IF EXISTS orders''')
        #Create table
        c.execute('''CREATE TABLE orders (MATERIAL text, COORDER text, DOC_NUMBER text, S_ORD_ITEM text,
            REFER_DOC text, REFER_ITM text)''')
        for tup in data:
            c.execute('INSERT INTO orders VALUES (?,?,?,?,?,?)', tup)
        conn.commit()
        conn.close()

    def detectDups(self, data):
        """
        Check if key has multiple values of value
        Input : data a tuple of tuples ((val1, val2,....),(......etc)
        Output : prints its results to terminal
        """
        db = {}
        for line in data:
            key = line[4] + line[5] #REFER_DOC & REFER_ITM
            value = line [0] #Material
            if key in db.keys():
                if value in db[key]: pass
                else : db[key].append(value)
            else: db[key] = [value]
        for key in db.keys():
            if len(db[key]) > 1: print (key, db[key])

if __name__ == '__main__':
    c = SAPFile()
    #c.fixLines1('ZO00077_P2W.txt','ZO00077_P2W_out.txt')
    data = c.getData('C:\Documents and Settings\\u104675\Desktop\Service Investigation\ZO00077_P2W_2014_RestrictedChars.txt')
    #c.loadToSqlite(data)
    c.detectDups(data)
