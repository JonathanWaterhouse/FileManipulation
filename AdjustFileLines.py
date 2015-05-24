import xlrd

__author__ = 'U104675'
class AdjustFileLines():
    def __init__(self,file,fileOut):
        f = open(file,'r')
        fout = open(fileOut,'w')
        self._lst, self._col_names = [], []
        i = 0
        for line in f:
            self._lst = line.strip().strip('"').strip('|').strip().split('|')
            self._lst = [el.strip() for el in self._lst]
            if i == 0:
                self._col_names = [el.strip() for el in self._lst]
                self._col_names.append('ZDESTCNTY') # A temp measure because I forgot to download ZREGION
            elif i>= 300000 and i <= 400000:
                self._fix_dec('ZAMNT_GC')
                self._fix_dec('ZAMNT_LC')
                self._fix_date('0CALDAY')
                self._fix_date('PSTNG_DATE')
                self._fix_calmonth('0CALMONTH')
                self._fix_calweek('0CALWEEK')
                self._fix_fiscper('0FISCPER')
                self._fix_date('ZPCKCFDT')
                self._fix_date('ZREVDATE')
                self._lst.append('0909') # A temp measure because I forgot to download ZREGION
            self
            if i>=300000 and i <= 400000:
                fout.write('|'.join(self._lst)+ '\n')
            i += 1
            #if i > 20: break

    def _fix_dec(self, fld_nm):
        idx = self._col_names.index(fld_nm)
        fld_val = self._lst[idx]
        if fld_val == '': return
        if fld_val[-1] == '-': fld_val = '-' + fld_val.rstrip('-')
        self._lst[idx] = fld_val.replace(',','')

    def _fix_date(self, fld_nm):
        idx = self._col_names.index(fld_nm)
        fld_val = self._lst[idx]
        if fld_val == '': return
        temp = fld_val[6:10] + fld_val[3:5] + fld_val[0:2]
        self._lst[idx] = "".join(temp)

    def _fix_calmonth(self, fld_nm):
        idx = self._col_names.index(fld_nm)
        fld_val = self._lst[idx]
        if fld_val == '': return
        temp = fld_val[4:6] + fld_val[0:4]
        self._lst[idx] = "".join(temp)

    def _fix_calweek(self, fld_nm):
        idx = self._col_names.index(fld_nm)
        fld_val = self._lst[idx]
        if fld_val == '': return
        temp = fld_val[4:6] + fld_val[0:4]
        self._lst[idx] = "".join(temp)

    def _fix_fiscper(self, fld_nm):
        idx = self._col_names.index(fld_nm)
        fld_val = self._lst[idx]
        if fld_val == '': return
        temp = fld_val[4:7] + fld_val[0:4]
        self._lst[idx] = "".join(temp)

class MergeFiles():
    def __init__(self, file,fileOut):
        fout = open(fileOut,'w')
        for fi in file:
            f = open(fi,'r')
            for line in f:
                fout.write(line+'\n')
            f.close()
        fout.close()

class XCYCTFCTR():
    def __init__(self, fi, fo):
        wb = xlrd.open_workbook(fi)
        ws = wb.sheet_by_name('ReasonCode')
        rows = range(ws.nrows)
        rc = {}
        for i in rows:
            if i > 0: rc[ws.cell_value(i,0)] = ws.cell_value(i,1)
        for k,v in rc.items() : print (k, v)

        ws = wb.sheet_by_name('Factor')
        rows = range(ws.nrows)
        cols = range(ws.ncols)
        factors = {}
        groups = []
        for i in rows:
            if i == 0:
                for j in cols:
                    if j > 0: groups.append(ws.cell_value(i,j))
                print (groups)
            else:
                for j in cols:
                    if j > 0:
                        try: factors[ws.cell_value(i,0)][groups[j-1]] = ws.cell_value(i,j)
                        except (KeyError):
                            factors[ws.cell_value(i,0)] = {}
                            factors[ws.cell_value(i,0)][groups[j-1]] = ws.cell_value(i,j)
        for k,v in sorted(factors.items()):
            print(k,v)

        #wb.close()
        fout = open(fo,'w')
        for code in sorted(rc.keys()):
            for BU in factors.keys():
                if rc[code] != 'NA': fout.write('0909' + '|' + BU + '|' + code + '|' + repr(factors[BU][rc[code]]) + '\n')
        fout.close()

if __name__ == '__main__' :
    ##folder = 'C:\Documents and Settings\\u104675\Desktop\Sales returns reserve\\'
    ##fi = folder + 'ZC00204P2W012014_OUT.txt'
    ##fo = folder + 'ZC00204.txt'
    ##adj = AdjustFileLines(fi, fo)
    folder = 'C:\Documents and Settings\\u104675\Desktop\Sales returns reserve\\'
    fi = folder + 'ReserveFactors.xlsx'
    fo = folder + 'XCYCTFCTR.txt'
    x = XCYCTFCTR(fi, fo)
    #fileIn = []
    #fileIn.append(folder + 'ZO00060_FEB_2013_0351_LOC.txt')
    #fileIn.append(folder + 'ZO00060_FEB_2013_0351_USD.txt')
    #fileIn.append(folder + 'ZO00060_FEB_2013_0286_LOC.txt')
    #fileIn.append(folder + 'ZO00060_FEB_2013_0286_USD.txt')
    #fileIn.append(folder + 'ZO00060_mar_ADJ.txt')
    #fileOut = folder + 'ZO00060_INPUT_DATA.txt'
    #c = MergeFiles(fileIn,fileOut)