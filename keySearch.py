import pandas as pd
import os

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
        #print(testFile,df.shape)
    except : 
        return (['-1',"Error : Cannnot Open Input Excel file...."])

    try:
        found = []
        colList = df.columns.tolist()
        for i in keyWords:
            if i in colList:
                found.append(i)
        if len(found) == 0:       
            return (['1',"Info : No Matching Columns found"]) 
        else:
            return (['1',"Found Columns: "+"".join(map(str, found))])  
    except : 
        return (['-1',"Error : Failed to find the columns"])               
                

def doContentSearch(testFile,keyWords,colName):
        
    pyPath, pyName = os.path.split(__file__)
    os.chdir(pyPath)
    try:
        df = pd.read_excel(testFile)
        #print(testFile,df.shape)
    except : 
        return (['-1',"Error : Cannnot Open Input Excel file...."])
    dfOut = df[df[str(colName)].str.contains('tai',case=False)]
    try:
        writer = pd.ExcelWriter('output.xlsx', engine='openpyxl')
        
        shape = []
        
        dummy = pd.DataFrame()
        dummy.to_excel(writer, sheet_name='Search Summary')
        dfOut = df[df[str(colName)].str.contains('tati',case=False)]        
        for kw in keyWords:
            dfOut = df[df[str(colName)].str.contains(kw,case=False)]
            tup = [kw,dfOut.shape[0]]
            #print(tup)
            shape.append(tup)
            dfOut.to_excel(writer, sheet_name=kw, index=False)
            
        summ = pd.DataFrame(shape,columns = ['Key-Word','Frequency'])
        summ.sort_values(by='Frequency',ascending=False).to_excel(writer, sheet_name='Search Summary', index=False)
        writer.save()
        return(['1',"Success! output.xlsx created with search results.."])
    except : 
        return (['-1',"Error : Failed to create output file. Please check if the file is already opened."]) 
    

if __name__ == '__main__':
    print("Please run this app Using 'app.py'.....")   