from tkinter import *
from tkinter import messagebox

import xlrd
import openpyxl


def pickcar():
    mywin=Tk()
    mywin.resizable(0,0)
    mywin.title("DEALER")
    
    mycolor="bisque2"
    mywin.config(bg=mycolor)

    def carpickup():    #window for carpickup
        window=Tk()
        mycolor="bisque2"
        window.config(bg=mycolor)
        loc='RentalDatabase.xlsx'
        wb=xlrd.open_workbook(loc)
        sheet=wb.sheet_by_name("Info abt renter")


        wb2=openpyxl.load_workbook("RentalDatabase.xlsx")
        sheet1=wb2["Renter history"]
        window.title("Car pickup")
        window.resizable(0,0)


        myfont=("Helveica","16")
        name=Label(window,text="Renters name:",font=myfont,bg=mycolor).grid(row=1,column=1,padx=5,pady=5)
        name_e=Entry(window)
        name_e.grid(row=1,column=2,padx=5,pady=5)

        phone=Label(window,text="Phone number:",font=myfont,bg=mycolor).grid(row=2,column=1,padx=5,pady=5)
        phone_e=Entry(window)
        phone_e.grid(row=2,column=2,padx=5,pady=5)

        


        def onclick():
            count=0
            for i in range(1,sheet.nrows):
                
                if name_e.get()==sheet.cell_value(i,2) and phone_e.get()==sheet.cell_value(i,3):
                    wbpxl=openpyxl.load_workbook("RentalDatabase.xlsx")
                    sheetx=wbpxl["All cars"]
                    sheet22=wb.sheet_by_name("All cars")
                    for j in range(1,sheet22.nrows):
                        if sheet.cell_value(i,5)==sheet22.cell_value(j,2):
                            ax="F"+str(j+1)
                            sheetx[ax]="NO"
                            wbpxl.save("RentalDatabase.xlsx")
                            messagebox.showinfo("Done","Done!")
                elif name_e.get()=="" or phone_e.get()=="":
                    messagebox.showerror("Error","You must fill all the boxes.")
                    break
                    
                else:
                    count+=1
            if count==sheet.nrows-1:
                messagebox.showerror("Error","User doesnt exist.")
                            
                  
            


        bt=Button(window,text="Finish",width=20,command=onclick,bg="grey").grid(row=4,column=2,padx=5,pady=5)

        window.mainloop()

    myfon=("Helvetica","12")

    label=Label(mywin,text="User name:",font=myfon,bg=mycolor)
    label.grid(row=1,column=1,padx=5,pady=5)

    label_e=Entry(mywin)
    label_e.grid(row=1,column=2,padx=5,pady=5)

    pword=Label(mywin,text="Password:",font=myfon,bg=mycolor)
    pword.grid(row=2,column=1,padx=5,pady=5)

    pword_e=Entry(mywin)
    pword_e.grid(row=2,column=2,padx=5,pady=5)


    def pickit():
        if label_e.get()=="dealer" and pword_e.get()=="iamthedealer":
            carpickup()
            
        else:
            messagebox.showerror("Error","Check username and password")
            



    b=Button(mywin,text="Confirm",command=pickit,font=("Helvetica","10","bold"),bg="grey").grid(row=3,columnspan=3,padx=10,pady=10)

    mywin.mainloop()







