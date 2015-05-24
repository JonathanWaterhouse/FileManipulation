__author__ = 'U104675'

class RemoveFileCols():
    def __init__(self, fin, fout):
        fi = open(fin,'r', encoding='Latin-1')
        fo = open(fout,'w')
        for line in fi:
            if line[0] == '|':
                line_list = line.split('|')
                del line_list[0]
                del line_list[len(line_list) - 1]
                line_list_strip = [el.strip() for el in line_list] # Remove spaces at start and end of values
                fo.write(','.join(line_list_strip) + '\n')

if __name__ == "__main__":
    path = 'C:\\Users\\u104675\\Desktop\\'
    filein = path + '0Customer_short.txt'
    fileout = path + '0Customer_short.csv'
    new_file = RemoveFileCols(filein, fileout)