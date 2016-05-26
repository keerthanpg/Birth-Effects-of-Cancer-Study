'''import pandas as pd
my_dic = pd.read_excel('chemicals.xlsx', index_col=0).to_dict()

for [k,v] in my_dic:
    print k, v
    break'''

from xlrd import open_workbook

book_C = open_workbook('chemicals.xlsx')
book_RT = open_workbook('chemicalsRT.xlsx')
book_DT = open_workbook('chemicalsDT.xlsx')

sheet_C = book_C.sheet_by_index(0)
sheet_RT = book_RT.sheet_by_index(0)
sheet_DT = book_DT.sheet_by_index(0)

# read header values into the list    
keys = [sheet_C.cell(0, col_index).value for col_index in xrange(sheet_C.ncols)]

ScorecardList = []
RT_list = []
DT_list = []
for row_index in xrange(1, sheet_C.nrows):
    d = {keys[col_index]: sheet_C.cell(row_index, col_index).value 
         for col_index in xrange(sheet_C.ncols)}
    d['Effect']='Cancer'
    ScorecardList.append(d)
print len(ScorecardList)
i=0
for row_index in xrange(1, sheet_RT.nrows):
    d = {keys[col_index]: sheet_RT.cell(row_index, col_index).value 
         for col_index in xrange(sheet_C.ncols)}
    i=i+1
    print i
    for chemical in ScorecardList[:]:
        if chemical['CAS No']==d['CAS No']:
            print ('CAS Matched %s', chemical['CAS No'] )
            chemical['Effect']=chemical['Effect'] + ', Reproductive Toxicity'
        else:
            ScorecardList.append(d)
'''
for row_index in xrange(1, sheet_DT.nrows):
    d = {keys[col_index]: sheet_DT.cell(row_index, col_index).value 
         for col_index in xrange(sheet_C.ncols)}
    for chemical in ScorecardList:
        if chemical['CAS No']==d['CAS No']:
            print ('CAS Matched')
            chemical['Effect']=chemical['Effect'] + ', Reproductive Toxicity'
        else:
            ScorecardList.append(d)'''
    

print len(ScorecardList)
