from tkinter import *
from tkinter import filedialog
from PIL import ImageTk, Image
import pandas as pd

import os
pyPath, pyName = os.path.split(__file__)
os.chdir(pyPath)

import keySearch as ks

def mySubmit():
    
    fileName=dfPath
    keyWordString=str(t2.get("1.0",END)).rstrip()
    keyWords=list(str(t2.get("1.0",END)).rstrip().split(','))
    colName=str(t1.get("1.0",END)).rstrip()
    
    writepath = 'logSearch.log'

    mode = 'a' if os.path.exists(writepath) else 'w'
    with open(writepath, mode) as f:
        f.write("\nFile Name: " + str(fileName))
        f.write("\nColumn Name: " +  colName)
        f.write("\nFile Name: " +  keyWordString)
        f.write("\n\n")

    Label(background='white').grid()

    if var.get() == 1:
        search = ks.doContentSearch(fileName,keyWords,colName)
        le = Label(root, text=str(search[1]))
        le.grid()
    else:
        search = ks.doColSearch(fileName,keyWords,colName)
        le = Label(root, text=str(search[1]))
        le.grid()           
    
def getFile():
    currdir = os.getcwd()
    Label(background='white').grid()
    try:
        file = filedialog.askopenfilename(parent=root, initialdir=currdir, title='Please select a directory')
      
        global dfPath
        dfPath = file
        df=pd.read_excel(file)
    
        le = Label(root, text=str(dfPath)+"\nLoaded with dimensions: "+str(df.shape), background='white')
        le.grid(row=20,column=1, sticky=NW)
    except:     
        le = Label(root, background='white', text="Error: File Selection Failed! Check File type. \n"+str('\n'.join(map(str, sys.exc_info()))))
        le.grid(row=20,column=1, sticky=NW)
  

#Define the Main Screen        
root = Tk()
root.title("Excel Scrapper")
#root.geometry("1000x600")
root.configure(background='white')

#Insert Logo
load = Image.open("py.jpg").resize((80, 55), Image.ANTIALIAS)
render = ImageTk.PhotoImage(load)
img = Label(root, image=render,background='white')
img.image = render
img.grid(row=0,column=0,sticky=NW)

#Lable: Col Name
l1 = Label(root, text="PE Analytics         ",background='white', font=('Segoe UI', 18, 'bold'), fg='grey')
l1.grid(row=0,column=3, sticky=E, columnspan=3)

Label(root, text=" ",background='white').grid()

#Text: Directory
t1 = Text(root, height = 1, width = len(str(os.getcwd()))+5, bd=2)
t1.insert(index=END,chars="Present Working Directory is "+str(os.getcwd()) )
t1.configure(state='disabled')
t1.grid(row=10,column=0, sticky=NW, columnspan=3)

Label(root, text=" ",background='white').grid()

b1 = Button(root, text='Choose File to Search', height="1", width="20", command=getFile)
b1.grid(row=20,column=0, sticky=NW)

Label(root, text=" ",background='white').grid()
Label(root, text=" ",background='white').grid()

#Lable: Col Name
l1 = Label(root, text="Column to be searched:",background='white')
l1.grid(row=30,column=0, sticky=NW)

#Text: Key-Words
t1 = Text(root, height = 1, width = 50, bd=2, bg='#D4D6D2')
t1.grid(row=40,column=0, sticky=NW, columnspan=3)

Label(root, text=" ",background='white').grid()

#Lable: Key-Words
l1 = Label(root, text="Key-Words to search (Comma Separated):",background='white')
l1.grid(row=50, sticky=NW)

#Text: Key-Words
t2 = Text(root, height = 1, width = 50, bd=2, bg='#D4D6D2')
t2.grid(row=60, sticky=NW, columnspan=3)

Label(root, text=" ",background='white').grid()

#Lable: Function
l1 = Label(root, text="Search Domain:", justify = LEFT,background='white')
l1.grid(row=70, sticky=NW)


#Radiobutton: Function
var = IntVar()
r1 = Radiobutton(root, text='Content in the Specified Column', variable=var, value=1, background='white') 
r2 = Radiobutton(root, text='Column Names in the Excel', variable=var, value=2, background='white')
r1.grid(row=80, sticky=NW)
r2.grid(row=90, sticky=NW)

Label(background='white').grid()

#Add Buttons
b2 = Button(root, text='Submit', height="2", width="30", command=mySubmit)
b2.grid(row=100, sticky=NW)

#grid(column = 0, row = 0)

if __name__ == '__main__':    
    root.mainloop()