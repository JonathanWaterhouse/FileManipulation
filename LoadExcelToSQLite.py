__author__ = 'U104675'
import sqlite3
import xlrd

class ExcelInSQL():
    def __init__(self, database, excel_file, sheet):
        workbook = xlrd.open_workbook(excel_file)
        worksheet = workbook.sheet_by_name(sheet)
        num_rows = worksheet.nrows -1
        num_cols = worksheet.ncols -1
        conn = sqlite3.connect(database)
        c = conn.cursor()

        #Delete table if already exists
        c.execute('DROP TABLE IF EXISTS ' + sheet)

        curr_row = -1
        while curr_row < num_rows:
            curr_row += 1
            row = worksheet.row(curr_row)
            #if curr_row == 2: break
            row_list = []
            curr_cell = -1
            while curr_cell < num_cols:
                curr_cell +=1
                cell_value = worksheet.cell_value(curr_row,curr_cell)
                cell_type = worksheet.cell_type(curr_row,curr_cell)
                row_list.append((cell_value, cell_type, curr_cell))
            if curr_row == 0:
                col_def = [col_names_reformat(el) + ' text' for el in row_list]
                head_def = [col_names_reformat(el) for el in row_list]
                sql_stmt = "CREATE TABLE " + sheet + " (" +  ", ".join(col_def) +")"
                print(sql_stmt)
                c.execute(sql_stmt)
            else:
                sql_stmt = "INSERT INTO " + sheet + " (" + ",".join(head_def) + ") VALUES (" + ",".join([repr(el[0]) for el in row_list]) +")"
                c.execute(sql_stmt)

        conn.commit()

def col_names_reformat(intup):
    # Cell Types: 0=Empty, 1=Text, 2=Number, 3=Date, 4=Boolean, 5=Error, 6=Blank.
    el = ("".join(x for x in intup[0] if x.isalnum()), intup[1], intup[2])
    if el[1] == 0 or el[0] == '' : ell = 'TEXT' + repr(el[2])
    else: ell = el[0]
    return ell

if __name__ == "__main__":
    path = 'C:\\Users\\u104675\\Desktop\\Open_Hub_Conversions\\'
    database = path + 'AAG.db'
    excel = path + "ZOHSA005.xlsx"
    sheet = 'ZOHSA005'
    table = ExcelInSQL(database, excel, sheet)
