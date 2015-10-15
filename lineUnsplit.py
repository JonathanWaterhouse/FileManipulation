import xlsxwriter
def Unsplit(fileIn, fileOut, field_sep):
    """
    takes a file downloaded by ZSE16 in unconverted format, where the lines
    are long enough to have wrapped. Joins the two have lines
    """
    fi = open(fileIn,'r')
    fo = open(fileOut,'w')
    inCount = 0
    outCount = 0
    l1 = []
    l2 = []
    for line in fi:
        #if line[0] != "|": continue
        if line[0] == "-":
            continue
        if inCount%2 == 0:
            ln = line.rstrip(field_sep + '\n')
            l1 = ln.split(field_sep)
        else:
            ln = line.rstrip('\n')
            l2 = ln.split(field_sep)
            l1.extend(l2)
            fo.write(field_sep.join(l1) + '|\n')
            l1, l2 = [], []
            outCount += 1
        inCount += 1
    fi.close()
    fo.close()

def tidy(fileIn, fileOut):
    """
    Tidy a file of lines extracted from ZSE16 in unconverted format.
    Remove leading and trailing '|', remove trailing '\n' and from each field
    on the line remove leading and trailing blanks
    """
    fi = open(fileIn,'r')
    fo = open(fileOut,'w')
    for line in fi:
        ln = line.lstrip('|').lstrip(' ').rstrip('|\n')
        fo.write(ln + '\n')
    fi.close()
    fo.close()

def wrtToExcel(fileIn, fileOut):
    """
    use xlsxwriter module to write an input text file with fields separated by
    '|' to excel.
    """
    fi = open(fileIn,'r')
    workbook = xlsxwriter.Workbook(fileOut)
    worksheet = workbook.add_worksheet()
    chunk = []
    i,j = 0, 0
    for line in fi:
        chunk = line.split('|')
        for el in chunk:
            el.strip(' ')
            worksheet.write(i,j,el)
            j += 1
        j = 0
        i += 1

if __name__ == '__main__':
    folder = 'C:\\Documents and Settings\\u104675\\Desktop\\'
    fileIn = folder + 'CRMContractLineItem_PSA.TXT'
    fileOut = folder + 'CRMContractLineItem_PSA_Unsplit.TXT'
    field_sep = '|'
    Unsplit(fileIn, fileOut, field_sep)

    fileIn = folder + 'ZO00217Jul1st-20th_UNSPLIT.txt'
    fileOut = folder + 'ZO00217Jul1st-20th_TIDY.txt'
    #tidy(fileIn, fileOut)

    fileIn = folder + 'ZO00217Jul1st-20th_TIDY.txt'
    fileOut = folder + 'ZO00217Jul1st-20th.xlsx'
    #wrtToExcel(fileIn, fileOut)