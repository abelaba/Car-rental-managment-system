from tkinter import *
from tkinter import messagebox

import xlrd
import openpyxl


def returncar():
    mywin=Tk()
    mywin.resizable(0,0)
    mywin.title("DEALER")

    mycolor="bisque2"
    mywin.config(bg=mycolor)

    def carreturnwindow():
        window=Tk()
        mycolor="bisque2"
        window.config(bg=mycolor)
        loc=('RentalDatabase.xlsx')
        wb=xlrd.open_workbook(loc)
        sheet=wb.sheet_by_name("Info abt renter")


        wb2=openpyxl.load_workbook("RentalDatabase.xlsx")
        sheet1=wb2["Renter history"]
        window.title("Car return")
        window.resizable(0,0)


        myfont=("Helveica","16")
        name=Label(window,text="Renters name:",font=myfont,bg=mycolor).grid(row=1,column=1,padx=5,pady=5)
        name_e=Entry(window)
        name_e.grid(row=1,column=2,padx=5,pady=5)

        plate=Label(window,text="Lisence plate:",font=myfont,bg=mycolor).grid(row=2,column=1,padx=5,pady=5)
        plate_e=Entry(window)
        plate_e.grid(row=2,column=2,padx=5,pady=5)

        car=Label(window,text="Vehicle name:",font=myfont,bg=mycolor).grid(row=3,column=1,padx=5,pady=5)
        car_e=Entry(window)
        car_e.grid(row=3,column=2,padx=5,pady=5)


        def onclick():        # to add on info abt renter sheet
            for i in range(1,sheet.nrows):
                if name_e.get()==sheet.cell_value(i,2) and plate_e.get()==sheet.cell_value(i,5) and car_e.get()==sheet.cell_value(i,0):
                    loc=('RentalDatabase.xlsx')
                    wb=xlrd.open_workbook(loc)
                    sheet3=wb.sheet_by_name("Renter history")
                    sheet2=wb2["Info abt renter"]
                    c=sheet3.nrows+1
                    a=65
                    for j in range(0,10):
                        character=chr(a)+str(c)
                        sheet1[str(character)]=sheet.cell_value(i,j)
                        a+=1
                    
                    
                    sheetx=wb2["All cars"]
                    sheet22=wb.sheet_by_name("All cars")
                    for k in range(1,sheet22.nrows):             # to add yes on all cars sheet
                        
                        if sheet.cell_value(i,5)==sheet22.cell_value(k,2):
                            
                            ax="F"+str(k+1)
                            sheetx[ax]="YES"
                            
                    sheet2.delete_rows(i+1,1)        
                    wb2.save("RentalDatabase.xlsx")       
                    messagebox.showinfo("Done","You are done.")
                    
                    window.destroy()
                    
                   
                    print("Done")
                    break
                elif name_e.get()=="" or plate_e.get()=="" or car_e.get()=="":
                    messagebox.showerror("Error","You must fill all the boxes.")
                    break


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


    def com():     # to check if the password and username are correct
        if label_e.get()=="dealer" and pword_e.get()=="iamthedealer":
            carreturnwindow()
            
        else:
            messagebox.showerror("Error","Check username and password")
            



    b=Button(mywin,text="Confirm",command=com,font=("Helvetica","10","bold"),bg="grey").grid(row=3,columnspan=3,padx=10,pady=10)

    mywin.mainloop()







