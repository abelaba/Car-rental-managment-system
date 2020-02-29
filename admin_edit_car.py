import openpyxl
import tkinter
from tkinter.ttk import *
from tkinter import ttk
from tkinter import messagebox


import xlrd

def editcar():

    mywin=tkinter.Tk()
    mywin.resizable(0,0)
    mywin.title("ADMIN")
    mycolor="bisque2"
    mywin.config(bg=mycolor)


    def returnn():        #editcarwindow
        loc=('RentalDatabase.xlsx')
        wb=xlrd.open_workbook(loc)
        sheet=wb.sheet_by_name("All cars")

        mycolor="bisque2"











        list1=[]                #list of cartypes
        for i in range(1,sheet.nrows):
            list1.append(sheet.cell_value(i,1))


            


        window=tkinter.Tk()

        window.config(bg=mycolor)
        window.resizable(0,0)
        window.title("Edit car")
        fon=("Arial Bold",)


        totalcars=tkinter.Label(window,text="Total cars = "+str(sheet.nrows-1),bg=mycolor,fg="green",font=fon).grid(row=1,column=1,padx=7,pady=7)
        
        frame1=tkinter.Frame(window,bg=mycolor)
        frame1.grid(row=2,column=1,padx=10,pady=10,ipadx=5,ipady=5)



        ctype=tkinter.Label(frame1,text="Choose car:",font=fon,bg=mycolor).grid(row=1,column=1,padx=5,pady=2)
        combo=Combobox(frame1,values=list1,width=20)
        combo.grid(row=1,column=2)








        myfont=("Helvetica","16")

        list2=["-","-","-"]
        def click():                                              #what button does when clickd
            temp1=tkinter.Label(frame1,width=25,height=2,bg=mycolor).grid(row=2,column=2,padx=5,pady=5)   # blank label
            temp2=tkinter.Label(frame1,width=25,height=2,bg=mycolor).grid(row=3,column=2,padx=5,pady=5)   # blank label
            temp3=tkinter.Label(frame1,width=25,height=2,bg=mycolor).grid(row=4,column=2,padx=5,pady=5)   # blank label
            list2.clear()
            for i in range(1,sheet.nrows):
                if combo.get()==sheet.cell_value(i,1):
                    list2.append(sheet.cell_value(i,2))
                    lisp2=tkinter.Label(frame1,text=list2[0],bg=mycolor,font=myfont).grid(row=2,column=2,padx=5,pady=5)
                    list2.append(sheet.cell_value(i,3))
                    milage2=tkinter.Label(frame1,text=list2[1],bg=mycolor,font=myfont).grid(row=3,column=2,padx=5,pady=5)
                    list2.append(sheet.cell_value(i,4))
                    price2=tkinter.Label(frame1,text=list2[2],bg=mycolor,font=myfont).grid(row=4,column=2,padx=5,pady=5)
                    
            

                    

            
        lisp1=tkinter.Label(frame1,text="License plate:",font=fon,bg=mycolor).grid(row=2,column=1,padx=5,pady=5)
        lisp2=tkinter.Label(frame1,text=list2[0],bg=mycolor,height=2).grid(row=2,column=2,padx=5,pady=5)






            
            




        milage1=tkinter.Label(frame1,text="Milage(KM):",font=fon,bg=mycolor).grid(row=3,column=1,padx=5,pady=5)
        milage2=tkinter.Label(frame1,text=list2[1],bg=mycolor,height=2)
        milage2.grid(row=3,column=2,padx=5,pady=5)

        price1=tkinter.Label(frame1,text="Price per day(birr):",font=fon,bg=mycolor).grid(row=4,column=1)
        price2=tkinter.Label(frame1,text=list2[2],bg=mycolor,height=2).grid(row=4,column=2,padx=5,pady=5)





        bt=tkinter.Button(frame1,text="Select",command=click,bg=mycolor,width=10).grid(row=1,column=3,padx=5,pady=5)

        ##############################################################################################################################







        #1st column

        umilage=tkinter.Label(frame1,text="Update milage:",font=fon,bg=mycolor).grid(row=5,column=1,padx=5,pady=5)
        uprice=tkinter.Label(frame1,text="Update price:",font=fon,bg=mycolor).grid(row=6,column=1,padx=5,pady=5)


            
        #2nd column


        umilage_e=Entry(frame1,width=25)
        umilage_e.grid(row=5,column=2,padx=5,pady=5)

        uprice_e=Entry(frame1,width=25)
        uprice_e.grid(row=6,column=2,padx=5,pady=5)



        wbpyxl=openpyxl.load_workbook("RentalDatabase.xlsx")

        sheet1=wbpyxl["All cars"]


            
        def error():
            if umilage_e.get()=="" or uprice_e.get()=="":
                messagebox.showwarning("Warning","You must fill out all the boxes.")
            
            else:
                editc()



                
        def editc():
            for i in range(1,sheet.nrows):
                if combo.get()==sheet.cell_value(i,1) and list2[0]==sheet.cell_value(i,2) :
                    mil="D"+str(i+1)
                    cash="E"+str(i+1)
                    sheet1[mil]=int(umilage_e.get())
                    sheet1[cash]=int(uprice_e.get())
                    wbpyxl.save("RentalDatabase.xlsx")
                    

            
        
            
        def delete():
            Msg=messagebox.askquestion("Done","Are you sure you want to delete this car?")
            if Msg=="yes":
                for i in range(1,sheet.nrows):
                    if combo.get()==sheet.cell_value(i,1) and list2[0]==sheet.cell_value(i,2) :
                        sheet1.delete_rows(i+1,1)
                        wbpyxl.save("RentalDatabase.xlsx")
                        messagebox.showinfo("Done","You have successfuly deleted the car")
                
        

        but=tkinter.Button(window,text="Edit car",command=error,bg="grey",width=64,height=2).grid(row=3,column=1)
        deletebutton=tkinter.Button(window,text="Delete",command=delete,bg="grey",width=64,height=2).grid(row=4,column=1)


        window.mainloop()

    myfon=("Helvetica","12")

    label=tkinter.Label(mywin,text="User name:",font=myfon,bg=mycolor)
    label.grid(row=1,column=1,padx=5,pady=5)

    label_e=tkinter.Entry(mywin)
    label_e.grid(row=1,column=2,padx=5,pady=5)

    pword=tkinter.Label(mywin,text="Password:",font=myfon,bg=mycolor)
    pword.grid(row=2,column=1,padx=5,pady=5)

    pword_e=tkinter.Entry(mywin)
    pword_e.grid(row=2,column=2,padx=5,pady=5)


    def com():
        if label_e.get()=="admin" and pword_e.get()=="iamtheadmin":
            returnn()
            
        else:
            messagebox.showerror("Error","Check username and password")
            



    b=tkinter.Button(mywin,text="Confirm",command=com,font=("Helvetica","10","bold"),bg="grey").grid(row=3,columnspan=3,padx=10,pady=10)

    mywin.mainloop()
