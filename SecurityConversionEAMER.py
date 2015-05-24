import xlrd

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
        while curr_row < num_rows:
            curr_row += 1
            row = wsht.row(curr_row)
            if row[8].value == 'EAMR' or row[8].value == 'MEAF IND SEA': cols_req = [1,0,7,9] # The columns we require in the order we need them
            else: cols_req = [1,0,7,8]
            data.append([el.value for el in [row[i] for i in cols_req]])

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
    excel_in = path + "SA Keith Bowyer To Approve - Keiths Updates.xlsx"
    excel_out = path + "SA EAMER Approved Gordy Format.xlsx"
    sheet = 'People'
    table = Excel(excel_in, excel_out, sheet)