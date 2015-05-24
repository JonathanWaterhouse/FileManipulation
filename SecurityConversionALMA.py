__author__ = 'U104675'
import xlrd
import xlsxwriter

class Excel():
    def __init__(self, excel_in, excel_out, sheet):
        #Grab excel data
        wbk = xlrd.open_workbook(excel_in)
        wsht = wbk.sheet_by_name(sheet)
        num_rows = wsht.nrows -1
        curr_row = -1
        data = []
        GARR = ['0111','0136','0241','0291','0408','0412','0417','0420','0421','0422','0427',
                '0428','0434','0456','0467','0750','0805','0C01','0C02','0C05']
        segments = {'Software & Solutions' : 'SG08' , 'Consumer & Film' : 'SG01' , 'PI&DI' : 'SG02' ,
                    'Print Systems' : 'SG06' , 'Micro 3D Printing & Packaging' : 'SG04' ,
                    'Enterprise Inkjet Systems' : 'SG07'}
        ALMA_rows = 0
        while curr_row < num_rows:
            curr_row += 1
            row = wsht.row(curr_row)
            cols_req = [1,0,8,9]
            #Select only ALMA rows
            if row[7].value != 'Gan, Simon': continue
            #Remove Alaris people
            if row[3].value == 'Kodak PI/DI Div': continue
            elif row[12].value == 'Remove':
                data.append([row[1].value, row[0].value, 'REMOVE ACCESS', 'REMOVE ACCESS'])
                ALMA_rows += 1
                continue
            else: ALMA_rows += 1
            # Substitute salesorgs for GARR
            sorg = row[9].value
            if sorg.find('GARR') != -1:
                row[9].value = sorg.replace('GARR' , ';'.join(GARR))
            #Replace text division name with segment key from dictionary stored earlier
            try:
                row[8].value = segments[row[8].value]
            except (KeyError): pass
            data.append([el.value for el in [row[i] for i in cols_req]])
        print('Number of ALMA rows selected for processing : ' + repr(ALMA_rows))


        #Reformat excel data
        i = 0
        div_col = 2 # col in internal table where division is
        reg_col = 3 # col in internal table where division is
        out = [] # Reformatted data, a list of lists
        for row in data:
            #Split divisions into separate lines
            if row[div_col].find(";") != -1: # ';' is found ie more than one element
                for div in row[div_col].split(';'):
                    out.append([row[0], row[1], div, 'Product','Segment'])
            else:
                out.append([row[0], row[1], row[div_col], 'Product', 'Segment'])
            #split regions into separate lines
            if row[reg_col].find(';') != -1: # ';' is found ie more than one element
                for reg in row[reg_col].split(';'):
                    out.append([row[0], row[1], reg, 'Geography','Sales Org'])
            else:
                if row[reg_col] == 'All geographies': category = 'All Geographies'
                else: category = 'Sales Org'
                out.append([row[0], row[1], row[reg_col],'Geography', category])
            i += 1

        #Output to new excel file
        wbko = xlsxwriter.Workbook(excel_out)
        wshto = wbko.add_worksheet(sheet)
        headings = ['User ID', 'User Name', 'Level', 'Type', 'Category']
        j = 0
        for h in headings:
            wshto.write(0,j,h)
            j += 1
        i = 1
        for row in out:
            j = 0
            for col in row:
                wshto.write(i,j,col)
                j += 1
            i += 1
        wbko.close()
        for el in out: print(el)

if __name__ == "__main__":
    path = 'C:\\Users\\u104675\\OneDrive - Eastman Kodak Company\\Sales Analysis Design Steward\\Reorganisation\\'
    excel_in = path + "SA Simon Gan To Approve - Simons Update.xlsx"
    excel_out = path + "SA ALMA Approved Gordy Format.xlsx"
    sheet = 'SA Users & Approvers'
    table = Excel(excel_in, excel_out, sheet)