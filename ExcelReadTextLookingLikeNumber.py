__author__ = 'U104675'
import xlrd
class Excel():
    def __init__(self, excel_file, sheet):
        wbk = xlrd.open_workbook(excel_file)
        wsht = wbk.sheet_by_name(sheet)
        num_rows = wsht.nrows -1
        num_cols = wsht.ncols -1
        curr_row = -1
        while curr_row < num_rows:
            curr_row += 1
            row = wsht.row(curr_row)
            row_list = []
            curr_cell = -1
            while curr_cell < num_cols:
                curr_cell +=1
                cell_value = wsht.cell_value(curr_row,curr_cell)
                cell_type = wsht.cell_type(curr_row,curr_cell)
                if curr_cell < 48 and cell_type == xlrd.XL_CELL_NUMBER: #Test for numeric values
                    row_list.append(repr(int(cell_value)))
                else:
                    row_list.append(cell_value)
        for el in row_list: print(el)

if __name__ == "__main__":
    path = 'C:\\Users\\u104675\\OneDrive - Eastman Kodak Company\\Sales Analysis Design Steward\\SA Cube Redimension\\'
    excel = path + "ZC00016_SRC_SMALL.xlsx"
    sheet = 'ZC00016_SRC'
    table = Excel(excel, sheet)
    