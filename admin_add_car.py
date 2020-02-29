from tkinter import *
from tkinter import messagebox

import xlrd

import openpyxl




def addcar():
    wb=xlrd.open_workbook("RentalDatabase.xlsx")
    wbpyxl=openpyxl.load_workbook("RentalDatabase.xlsx")

    sheet=wb.sheet_by_name("All cars")

    sheetp=wbpyxl["All cars"]


    window=Tk()
    mycolor="bisque2"
    window.config(bg=mycolor)
    window.resizable(0,0)
    window.title("Add car")

    myfont=("Helvetica","14","bold")

    ctype=Label(window,text="Car type:",font=myfont,bg=mycolor).grid(row=1,column=1,padx=5,pady=5)
    cmodel=Label(window,text="Car model:",font=myfont,bg=mycolor).grid(row=2,column=1,padx=5,pady=5)
    mile=Label(window,text="Milage:",font=myfont,bg=mycolor).grid(row=3,column=1,padx=5,pady=5)
    Lplate=Label(window,text="License plate:",font=myfont,bg=mycolor).grid(row=4,column=1,padx=5,pady=5)
    Price=Label(window,text="Price per day:",font=myfont,bg=mycolor).grid(row=5,column=1,padx=5,pady=5)


    ctype_e=Entry(window)
    ctype_e.grid(row=1,column=2,padx=5,pady=5)

    cmodel_e=Entry(window)
    cmodel_e.grid(row=2,column=2,padx=5,pady=5)

    mile_e=Entry(window)
    mile_e.grid(row=3,column=2,padx=5,pady=5)

    lplate_e=Entry(window)
    lplate_e.grid(row=4,column=2,padx=5,pady=5)

    price_e=Entry(window)
    price_e.grid(row=5,column=2,padx=5,pady=5)


    nextrow=sheet.nrows+1     #sheet=All cars
    cartype="A"+str(nextrow)
    carmodel="B"+str(nextrow)
    liplate="C"+str(nextrow)
    milage="D"+str(nextrow)
    priceperd="E"+str(nextrow)
    available="F"+str(nextrow)


    def click():
        sheetp[cartype]=ctype_e.get()
        sheetp[carmodel]=cmodel_e.get()
        sheetp[liplate]=lplate_e.get()
        sheetp[milage]=mile_e.get()
        sheetp[priceperd]=price_e.get()
        sheetp[available]="YES"
        window.destroy()
        wbpyxl.save("RentalDatabase.xlsx")
        messagebox.showinfo("Done","You have added a new car.")

    def onclick():
        count=0
        for i in range (1,sheet.nrows):    #for checking if car exits by license pate
            if lplate_e.get()==sheet.cell_value(i,2):
                count+=1
        if ctype_e.get()=="" or cmodel_e.get()=="" or lplate_e.get()=="" or mile_e.get()=="" or price_e.get()=="":
            messagebox.showwarning("WARNING","You must fill out all the boxes!")
        elif count!=0:
            messagebox.showwarning("WARNING","Car already exists!")
            
        else:
            box=messagebox.askquestion("?","Are you sure you want to add this car?")
            if box=="yes":
                click()
        

    bu=Button(window,text="Add car",width=40,height=2,bg="grey",command=onclick).grid(row=6,columnspan=4)


    window.mainloop()


window=Tk()
window.title("Admin")

mycolor="grey14"

backButton=Button(window,text="Back",width=7,bg="grey").grid(row=1,column=1,padx=5,pady=5,sticky="W")
addCarButton=Button(window,text="Add Car",bg="grey",width=30,height=20).grid(row=2,column=1,padx=17,pady=17)

editCarButton=Button(window,text="Edit Car",bg="grey",width=30,height=20).grid(row=2,column=2,padx=17,pady=17)




window.config(bg=mycolor)

addcar()

window.mainloop()











