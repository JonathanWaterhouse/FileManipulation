import collections
import xml.dom.minidom as mdom
import xlwt

__author__ = 'U104675'
#Initialisations
path = 'C:\\Users\\u104675\\Desktop\\'
output_rows = {}
output_cols = {}
output = {}
titles = {0:'Calls Offered', 1:'Calls Answered', 2:'Calls terminated', 3:'Calls Abandoned'}
dom = mdom.parse(path + 'Sep2014 CallsXML.xml')
#Get All row labels
# RowTotal elements have format like
# <RowTotal RowNumber="63">16:30</RowTotal>
# <RowTotal RowNumber="64">16:45</RowTotal>
# <RowTotal RowNumber="65">Credit_France - 994</RowTotal>
row_total_list = dom.getElementsByTagName('RowTotal')
times = []
for rt in row_total_list:
    row_nbr = rt.attributes['RowNumber'].value
    nodes = rt.childNodes
    for node in nodes:
        if node.nodeType == node.TEXT_NODE:
            # use an test for numeric on 1st two characters to see if it is a time or team title (MIGHT BREAK!)
            if node.data[0:2].isnumeric(): times.append((row_nbr, node.data))
            else: # At end of a section we get to the team name. Attach it to the times for the team
                team = node.data
                # Ensure we have row number for the team title row to match up with the cell below
                times.append((row_nbr,team))
                # The rows output is a dictionary { row_nbr : (time, team), .......}
                for t in times: output_rows[int(t[0])] = (t[1],team)
                times = []
print(output_rows)

#Get All Column labels
#Col Totals have a format like
#<ColumnTotal ColumnNumber="5">Fri</ColumnTotal>
#<ColumnTotal ColumnNumber="6">Sat</ColumnTotal>
#<ColumnTotal ColumnNumber="7">Total</ColumnTotal>
col_total_list = dom.getElementsByTagName('ColumnTotal')
for ct in col_total_list:
    col_nbr = ct.attributes['ColumnNumber'].value
    nodes = ct.childNodes
    for node in nodes:
        if node.nodeType == node.TEXT_NODE:
            # The cols output is of format { col_nbr : day, .....col_nbr : total}
            output_cols[int(col_nbr)] = node.data
print(output_cols)

#Get all data cells with row column identification
#Cells have values like this
#<Cell RowNumber="0" ColumnNumber="1">
#    <CellValue Index="0">
#      <FormattedValue>0</FormattedValue>
#      <Value>0.00</Value>
#    </CellValue>
#    <CellValue Index="1">
#       <FormattedValue>0</FormattedValue>
#       <Value>0.00</Value>
#    </CellValue>
cell_list = dom.getElementsByTagName('Cell')
for cl in cell_list:
    row_num = cl.attributes['RowNumber'].value
    col_num = cl.attributes['ColumnNumber'].value
    nodes = cl.getElementsByTagName('CellValue')
    if col_num == "7": #Totals Only
        for node in nodes:
            index = node.attributes['Index'].value
            child_nodes = node.getElementsByTagName('Value')
            for child_node in child_nodes:
                #print(child_node.nodeType)
                cc_nodes = child_node.childNodes
                for cc_node in cc_nodes:
                    if cc_node.nodeType == mdom.Node.TEXT_NODE:
                        #Store as{(row,col,type):value]
                        output[(int(row_num),int(col_num),int(index))] = cc_node.data
print(output)
od = collections.OrderedDict(sorted(output.items()))
outlist = []
for k,v in od.items():
    outlist.append([output_rows[k[0]][0],output_rows[k[0]][1], titles[k[2]], v])
for el in outlist: print (el)

# Fill in missing times