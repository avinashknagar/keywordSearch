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
    keyWords=list(str(t2.get("1.0",END)).rstrip().split(','))
    colName=str(t1.get("1.0",END)).rstrip()

    Label(background='white').pack(anchor=W)

    if var.get() == 1:
        search = ks.doContentSearch(fileName,keyWords,colName)
        le = Label(root, text=str(search[1]))
        le.pack(anchor=W)
    else:
        search = ks.doColSearch(fileName,keyWords,colName)
        le = Label(root, text=str(search[1]))
        le.pack(anchor=W)           
    
def getFile():
    currdir = os.getcwd()
    Label(background='white').pack(anchor=W)
    try:
        file = filedialog.askopenfilename(parent=root, initialdir=currdir, title='Please select a directory')
      
        global dfPath
        dfPath = file
        df=pd.read_excel(file)
    
        le = Label(root, text="File Selected: "+str(dfPath)+"\nLoaded with dimensions: "+str(df.shape))
        le.pack(anchor=W) 
    except:     
        le = Label(root, text="Error: File Selection Failed! Check File type. \n"+str('\n'.join(map(str, sys.exc_info()))))
        le.pack(anchor=W) 
  

#Define the Main Screen        
root = Tk()
root.title("Excel Scrapper")
root.geometry("700x600")
root.configure(background='white')

#Insert Logo
load = Image.open("py.jpg").resize((80, 55), Image.ANTIALIAS)
render = ImageTk.PhotoImage(load)
img = Label(root, image=render,background='white')
img.image = render
img.pack(anchor=W)

Label(background='white').pack(anchor=W)

#Text: Directory
t1 = Text(root, height = 1, width = 100, bd=0)
t1.insert(index=END,chars="Present Working Directory is "+str(os.getcwd()) )
t1.configure(state='disabled')
t1.pack(anchor=W)

Label(background='white').pack()

b1 = Button(root, text='Choose File to Search', height="1", width="20", command=getFile)
b1.pack(anchor=W)

Label(background='white').pack()

#Lable: Col Name
l1 = Label(root, text="Column to be searched:",background='white')
l1.pack(anchor=W)

#Text: Key-Words
t1 = Text(root, height = 1, width = 50, bd=2, bg='#D4D6D2')
t1.pack(anchor=W)

Label(background='white').pack()

#Lable: Key-Words
l1 = Label(root, text="Key-Words to search (Comma Separated):",background='white')
l1.pack(anchor=W)

#Text: Key-Words
t2 = Text(root, height = 1, width = 50, bd=2, bg='#D4D6D2')
t2.pack(anchor=W)

Label(background='white').pack(anchor=W)

#Lable: Function
l1 = Label(root, text="Search Domain:", justify = LEFT,background='white')
l1.pack(anchor=W)

#Radiobutton: Function
var = IntVar()
r1 = Radiobutton(root, text='Content in the Specified Column', variable=var, value=1, background='white') 
r2 = Radiobutton(root, text='Column Names in the Excel', variable=var, value=2, background='white')
r1.pack(anchor=W)
r2.pack(anchor=W)

Label(background='white').pack(anchor=W)


#Add Buttons
b2 = Button(root, text='Submit', height="2", width="30", command=mySubmit)
b2.pack(anchor=W)

#grid(column = 0, row = 0)

if __name__ == '__main__':    
    root.mainloop()