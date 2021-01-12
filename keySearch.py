import pandas as pd
import os
import re

def initInput():
    file = open('input.txt','r')     
    myInput = file.read().split('\n')
    #print("Searching below mentioned Key-Words in file '{}' ".format(myInput[0]))
    #print()
    return [myInput[0],myInput[1:]]

def doColSearch(testFile,keyWords,colName):  
    pyPath, pyName = os.path.split(__file__)
    os.chdir(pyPath)    
    
    keyWords.append(colName)
    try:
        df = pd.read_excel(testFile)
        df.columns = map(str.lower, df.columns)
    except : 
        return (['-1',"Error : Cannnot Open Input Excel file...."])

    try:
        found = []
        colList = df.columns.tolist()
        for i in keyWords:
            if str(i).lower() in list(map(str.lower,colList)):
                found.append(i)
        if len(found) == 0:   
            return (['1',"Info : No Matching Columns found"]) 
        else:
            return (['1',"Found Columns: "+", ".join(map(str, found))])  
    except : 
        return (['-1',"Error : Failed to find the columns"])               

def SearchList(kwList,string):
    s = ''
    for x in kwList:
        xLen = len(x.split())
        if xLen > 1:
            if (x.lower() in string.lower()):
                s = s+', '+str(x)
        elif xLen == 1:
            if (x.lower() in string.lower().split()):
                s = s+', '+str(x)
    return s[2:]        

def doContentSearch(testFile,keyWords,colName):
        
    pyPath, pyName = os.path.split(__file__)
    os.chdir(pyPath)
    try:
        df = pd.read_excel(testFile)
        df.columns = map(str.lower, df.columns)
        df = df[df[str(colName).lower()].notnull()]
    except : 
        return (['-1',"Error : Cannnot Open Input Excel file...."])
    try:
        writer = pd.ExcelWriter('output.xlsx', engine='openpyxl')
                
        dummy = pd.DataFrame()
        dummy.to_excel(writer, sheet_name='Search Summary')
        
        #Creating Master Sheet:
        dfMaster = df
        dfMaster['Master Search'] = df[str(colName).lower()].apply(lambda x: SearchList(keyWords,x))
        dfMaster['len'] = df['Master Search'].apply(lambda x: len(x))
        dfMaster['Number of Matches'] = df['Master Search'].apply(lambda x: 0 if len(x) == 0 else len(list(x.split(','))))
        
        dfMaster = dfMaster[dfMaster['len']>0]
        colList = [c for c in dfMaster if c not in ['Master Search', 'Number of Matches']]
        dfMaster = dfMaster[['Number of Matches','Master Search']+colList]
        dfMaster.to_excel(writer, sheet_name='Master Sheet', index=False)
        
        #Dict for all Sheets
        myDict = {}
        
        #List for Home Sheet
        shape = []
  
        #Rest of the Sheets
        for kw in keyWords:
            if len(kw.split()) > 1:
                dfOut = df[df[str(colName).lower()].str.contains(re.escape(kw),case=False)]
            else:
                dfOut = df[df[str(colName).lower()].apply(lambda x: kw.lower() in x.lower().split())]
                
            tup = [kw,dfOut.shape[0]]
            myDict[kw] = dfOut
            shape.append(tup)
            

        summ = pd.DataFrame(shape,columns = ['Key-Word','Frequency'])
        summ = summ.sort_values(by='Frequency',ascending=False)
        summ.reset_index(drop=True, inplace=True)

        #Print to Excel (Home Sheet)
        summ.to_excel(writer, sheet_name='Search Summary', index=True)

        #Print to Excel (Rest of the Sheets)        
        for x in range(summ.shape[0]):
            if summ['Frequency'][x] > 0:    
                dfOut = myDict[str(summ['Key-Word'][x])]
                dfOut.to_excel(writer, sheet_name=str(x), index=False)
        
        writer.save()
        return(['1',"Success! output.xlsx created with search results.."])
    except Exception as e: 
        return (['-1',"Error : Failed to create output file. Please check if the file is already opened :"+str(e)]) 
    

if __name__ == '__main__':
    #k = doContentSearch('C:/Users/avina/Documents/UpWork/keywordSearch/keywordSearch-master/DEALS.xls',['gosh!','future', 'looking to buy'],'target')
    #k = SearchList(['covid-19','looking to sell', 'sign'], 'signed the deal looking to selL')
    #print(k)
    print("Please run this app Using 'app.py'.....")   
