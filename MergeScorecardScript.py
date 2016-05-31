from xlrd import open_workbook
from sets import Set
import ast
from csv import DictWriter
import unicodedata
import json

book_C = open_workbook('chemicals.xlsx')
book_RT = open_workbook('chemicalsRT.xlsx')
book_DT = open_workbook('chemicalsDT.xlsx')

sheet_C = book_C.sheet_by_index(0)
sheet_RT = book_RT.sheet_by_index(0)
sheet_DT = book_DT.sheet_by_index(0)

# read header values into the list    
keys = [sheet_C.cell(0, col_index).value for col_index in xrange(sheet_C.ncols)]

ScorecardList = {}
RT_list = []
DT_list = []
for row_index in xrange(1, sheet_C.nrows):
    d = {keys[col_index]: sheet_C.cell(row_index, col_index).value 
         for col_index in xrange(sheet_C.ncols)}
    d['Effect']=str.encode('Cancer', 'utf-8')
    ScorecardList[d['CAS No']]=d
    #print ScorecardList[d['CAS No']]
#print len(ScorecardList)
i=0
for row_index in xrange(1, sheet_RT.nrows):
    d = {keys[col_index]: sheet_RT.cell(row_index, col_index).value 
         for col_index in xrange(sheet_C.ncols)}
    i=i+1
    #print i
    if d['CAS No']in ScorecardList:
        #print ScorecardList[d['CAS No']]
        #print ('CAS Matched for RT %s', d['CAS No'] )
        
        ScorecardList[d['CAS No']]['Effect']=ScorecardList[d['CAS No']]['Effect'] + str.encode(', Reproductive Toxicity', 'utf-8')

        #print ScorecardList[d['CAS No']]
    else:
        d['Effect']=str.encode('Reproductive Toxicity', 'utf-8')
        ScorecardList[d['CAS No']]=d
#print len(ScorecardList)

j=0
for row_index in xrange(1, sheet_DT.nrows):
    d = {keys[col_index]: sheet_DT.cell(row_index, col_index).value 
         for col_index in xrange(sheet_C.ncols)}
    j=j+1
    #print j
    if d['CAS No']in ScorecardList:
        #print ('CAS Matched for DT%s', d['CAS No'] )
        
        #print ScorecardList[d['CAS No']]
        ScorecardList[d['CAS No']]['Effect']=ScorecardList[d['CAS No']]['Effect'] + str.encode(', Developmental Toxicity', 'utf-8')
    else:
        d['Effect']=str.encode('Developmental Toxicity', 'utf-8')
        ScorecardList[d['CAS No']]=d
#print len(ScorecardList)
#print i
#print j

'''def convert(input):
    if isinstance(input, dict):
        return dict((convert(key), convert(value)) for key, value in input.iteritems())
    elif isinstance(input, list):
        return [convert(element) for element in input]
    elif isinstance(input, unicode):
        return input.encode('utf-8')
    else:
        return input

Scorecard_String=convert(ScorecardList)'''
Scorecard_List=[]
for [k,v] in ScorecardList.iteritems():
    Scorecard_List.append(v)
    
    
error=[]    

with open('Scorecard.csv','w') as outfile:
    writer = DictWriter(outfile, ('Chemical Name', 'CAS No','References', 'Effect'))
    writer.writeheader()
    for i in xrange(len(Scorecard_List)):        
        try:
            writer.writerow(Scorecard_List[i])
            
        except :            
            
            Scorecard_List[i]['Chemical Name']=unicodedata.normalize('NFKD', Scorecard_List[i]['Chemical Name']).encode('ascii','ignore')
                       
            try:
                writer.writerow(Scorecard_List[i])
            except Exception as ex:
                template = "An exception of type {0} occured. Arguments:\n{1!r}"
                message = template.format(type(ex).__name__, ex.args)
                print Scorecard_List[i] 
                print message
                error.append(Scorecard_List[i])

for i in xrange(len(error)):
    print('Add row to Scorecard:')
    print error[i]

outfile.close()

with open('Scorecard.txt','w') as outfile:
    json.dump(ScorecardList, outfile)


           
