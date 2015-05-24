__author__ = 'U104675'
class RemoveFileCols():
    def __init__(self, fin, fout):
        fi = open(fin,'r', encoding='Latin-1')
        fo = open(fout,'w')
        for line in fi:
            fo.write(line[0:58] + line [214:251] + '\n')

if __name__ == "__main__":
    path = 'C:\\Users\\u104675\\Desktop\\'
    filein = path + '0Customer.txt'
    fileout = path + '0Customer_short.txt'
    new_file = RemoveFileCols(filein, fileout)
