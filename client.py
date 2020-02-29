import openpyxl
import tkinter
from tkinter.ttk import *
from tkinter import ttk
from tkinter import messagebox
from datetime import date

import xlrd

def user(thelist,window2,frameForFirstPage):
    loc=('RentalDatabase.xlsx')
    wb=xlrd.open_workbook(loc)
    sheet=wb.sheet_by_name("All cars")

    mycolor="grey14"
    labelColor="AntiqueWhite3"











    list1=[]                #list of cartypes with information
    for i in range(1,sheet.nrows):
        list1.append(sheet.cell_value(i,1))


        


    
    window2.title("Welcome back,"+thelist[0])
    fon=("Arial Bold",)
   
    frameForFirstPage.grid_forget()

    frameForUser=tkinter.Frame(window2,bg=mycolor)
    frameForUser.grid(row=1,column=1)

    def back():
        window2.title("CaRent!")
        emptyMenu=tkinter.Menu(window2)
        window2.config(menu=emptyMenu)
        frameForFirstPage.grid(row=1,column=1)
        frameForUser.grid_forget()
    

    backButton=tkinter.Button(frameForUser,text="Back",command=back,width=7,bg="grey").grid(row=1,column=1,padx=5,pady=5,sticky="W")

    carsavailable=tkinter.Label(frameForUser,text="Cars available:",font=fon,bg=mycolor,fg=labelColor).grid(row=2,column=1,padx=5,pady=2)
    combo=Combobox(frameForUser,values=list1,width=20)
    combo.grid(row=2,column=2)





    myfont=("Helvetica","16")
    
    list2=["-","-","-"]
    def displayinfo():                                              #for displayinginfo
        temp1=tkinter.Label(frameForUser,width=25,height=2,bg=mycolor,fg=labelColor).grid(row=3,column=2,padx=5,pady=5)   # blank label
        temp2=tkinter.Label(frameForUser,width=25,height=2,bg=mycolor,fg=labelColor).grid(row=4,column=2,padx=5,pady=5)   # blank label
        temp3=tkinter.Label(frameForUser,width=25,height=2,bg=mycolor,fg=labelColor).grid(row=5,column=2,padx=5,pady=5)   # blank label
        list2.clear()
        for i in range(1,sheet.nrows):          #sheet= all cars
            if combo.get()==sheet.cell_value(i,1):
                
                list2.append(sheet.cell_value(i,2))
                lisp2=tkinter.Label(frameForUser,text=list2[0],bg=mycolor,font=myfont,fg=labelColor).grid(row=3,column=2,padx=5,pady=5)
                
                list2.append(sheet.cell_value(i,3))
                milage2=tkinter.Label(frameForUser,text=list2[1],bg=mycolor,font=myfont,fg=labelColor).grid(row=4,column=2,padx=5,pady=5)
                
                list2.append(sheet.cell_value(i,4))
                price2=tkinter.Label(frameForUser,text=list2[2],bg=mycolor,font=myfont,fg=labelColor).grid(row=5,column=2,padx=5,pady=5)
                
                list2.append(sheet.cell_value(i,5))
                
                if list2[3]!="YES":
                    shop=tkinter.Label(frameForUser,text="Car is not inside the shop.",bg=mycolor,fg="red").grid(row=7,column=2,padx=5,pady=5,ipadx=2)
                else:
                    shop=tkinter.Label(frameForUser,bg=mycolor,width=25).grid(row=7,column=2,padx=5,pady=5)



                
    bt=tkinter.Button(frameForUser,text="Select",command=displayinfo,bg=mycolor,width=10,fg=labelColor).grid(row=2,column=3,padx=5,pady=5)   #button for the combobox
        
    lisp1=tkinter.Label(frameForUser,text="License plate:",font=fon,bg=mycolor,fg=labelColor).grid(row=3,column=1,padx=5,pady=5)
    lisp2=tkinter.Label(frameForUser,text=list2[0],bg=mycolor,height=2,fg=labelColor).grid(row=3,column=2,padx=5,pady=5)

    shoptemp=tkinter.Label(frameForUser,bg=mycolor,width=25).grid(row=7,column=2,padx=5,pady=5)




        
        




    milage1=tkinter.Label(frameForUser,text="Milage(KM):",font=fon,bg=mycolor,fg=labelColor).grid(row=4,column=1,padx=5,pady=5)
    milage2=tkinter.Label(frameForUser,text=list2[1],bg=mycolor,height=2,fg=labelColor)
    milage2.grid(row=4,column=2,padx=5,pady=5)

    price1=tkinter.Label(frameForUser,text="Price per day(birr):",font=fon,bg=mycolor,fg=labelColor).grid(row=5,column=1)
    price2=tkinter.Label(frameForUser,text=list2[2],bg=mycolor,height=2,fg=labelColor).grid(row=5,column=2,padx=5,pady=5)

    drent=tkinter.Label(frameForUser,text="Number of days:",font=fon,bg=mycolor,fg=labelColor).grid(row=6,column=1,padx=5,pady=5)
    daysrented=Spinbox(frameForUser,from_=0,to=10000000)
    daysrented.grid(row=6,column=2,padx=5,pady=5)


    

    ##############################################################################################################################

    

    def changepass():      # for changing password on the menu
        wbpyxl=openpyxl.load_workbook("RentalDatabase.xlsx")
        sheeter=wbpyxl["Users"]

        mycolor="grey14"
        labelColor="AntiqueWhite3"
        
        window=tkinter.Tk()
        window.title("Change password")
        window.resizable(0,0)
        window.config(bg=mycolor)
        
        label1=tkinter.Label(window,text="Old password",bg=mycolor,fg=labelColor).grid(row=1,column=1,padx=5,pady=5)
        label1=tkinter.Label(window,text="New password",bg=mycolor,fg=labelColor).grid(row=2,column=1,padx=5,pady=5)

        entry_oldpassword=tkinter.Entry(window,bg=mycolor,fg=labelColor)
        entry_oldpassword.grid(row=1,column=2,padx=5,pady=5)
        entry_newpassword=tkinter.Entry(window,bg=mycolor,fg=labelColor)
        entry_newpassword.grid(row=2,column=2,padx=5,pady=5)
        
        def change2():
            
            if entry_oldpassword.get()==thelist[4]:
                if entry_newpassword.get()!="":
                    sheeter[thelist[5]]=entry_newpassword.get()
                    wbpyxl.save("RentalDatabase.xlsx")
                    messagebox.showinfo("Done","You have successfully changed your password.")
                    window.destroy()
                else:
                    messagebox.showwarning("Warning","You must fill out all the boxes")
            else:
                messagebox.showwarning("Warning","Make sure your old password is correct.")
                
        changebuttn=tkinter.Button(window,text="Change password",command=change2,bg="grey",width=33).grid(row=3,column=1,columnspan=6,padx=3,pady=3)

        window.mainloop()


        
    themenu=tkinter.Menu(window2)      #menu for client window
    themenu.add_command(label="Change password",command=changepass)
    window2.config(menu=themenu)
    


    wbpyxl=openpyxl.load_workbook("RentalDatabase.xlsx")

    sheet1=wbpyxl["Info abt renter"]

    

        
    def error():
        sheet_info=wb.sheet_by_name("Info abt renter")
        count=0
        for i in range(1,sheet_info.nrows):
            if thelist[2]==sheet_info.cell_value(i,3):
                count+=1
            
        if daysrented.get()=="":
            messagebox.showwarning("Warning","You must fill out all the boxes.")
        elif list2[3]!="YES":
            messagebox.showwarning("Warning","Car is not inside the shop,Choose another car.")
        elif count!=0:
            messagebox.showwarning("Warning","You can only rent one car at a time.")
            

        else:
            
            rentcar()
            



            

    def rentcar():
        sheetforshop=wbpyxl["All cars"]
        sheet_info=wb.sheet_by_name("Info abt renter")
        Msgbox=messagebox.askquestion("Done","Are you sure?")
        if Msgbox=="yes":
            i=sheet_info.nrows+1
            day=date.today()
            a="A"+str(i)    
            c="C"+str(i) #renters name      from reg
            d="D"+str(i) #phone number      from reg
            e="E"+str(i) #email             from reg
            f="F"+str(i) #license
            g="G"+str(i) #rented date
            h="H"+str(i) # days rented
            I="I"+str(i) #total price
            j="J"+str(i)
            sheet1[a]=combo.get()
            sheet1[c]=thelist[0]+" "+thelist[1]
            sheet1[d]=thelist[2]
            sheet1[e]=thelist[3]
            sheet1[f]=list2[0]
            sheet1[g]=day
            sheet1[h]=int(daysrented.get())
            sheet1[I]=float(int(daysrented.get())*int(list2[2]))
            sheet1[j]=day.month
            
            for i in range(1,sheet.nrows):
                if combo.get()==sheet.cell_value(i,1) and list2[0]==sheet.cell_value(i,2) :
                    pending="F"+str(i+1)
                    sheetforshop[pending]="Pending"

                    
            wbpyxl.save("RentalDatabase.xlsx")
            messagebox.showinfo("Done","You can now come and pick the car up at the shop")
            main.destroy()
            

    def mylicense():     # for license agrrement
        
        licens=open("license.txt","r")

        x=licens.read()

        main=tkinter.Tk()
        main.resizable(0,0)
        main.title("LICENSE AGREEMENT")


        window=tkinter.Frame(main)
        window.grid(row=1,column=1)

        y=tkinter.Scrollbar(window)
        y.pack(side="right",fill="y")

        txt=tkinter.Text(window,yscrollcommand=y.set,width=70,height=15)

        txt.pack(side="left")
        txt.insert("insert",x)


        y.config(command=txt.yview)


        window2=tkinter.Frame(main)
        window2.grid(row=2,column=1)

        def ifno():
            main.destroy()

        pay=str(int(daysrented.get())*int(list2[2]))+" BR"
        a1=tkinter.Label(window2,text="Total payement = "+pay,font=16,fg="red").grid(row=1,column=1,padx=5,pady=5)
        
        a2=tkinter.Label(window2,text="Do you agree with this terms?",font=16).grid(row=2,column=1,padx=5,pady=5)
        b1=tkinter.Button(window2,text="YES",width=15,command=error).grid(row=2,column=2,padx=5,pady=5)
        b2=tkinter.Button(window2,text="NO",width=15,command=ifno).grid(row=2,column=3,padx=5,pady=5)



        main.mainloop()
        
    rentcarbutton=tkinter.Button(frameForUser,text="RENT A CAR",command=mylicense,width=70,height=2,bg="grey").grid(row=8,column=1,columnspan=8)





    
 
