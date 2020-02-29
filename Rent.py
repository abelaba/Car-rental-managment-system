from tkinter import *
from tkinter import messagebox
import xlrd
import openpyxl

from client import user

#from admin_add_car import addcar
from admin_edit_car import editcar

from dealer_pick_up_car import pickcar
from dealer_return_car import returncar

window2=Tk()
mycolor="grey14"
labelColor="AntiqueWhite3"
window2.resizable(0,0)
window2.title("CaRent!")
window2.config(bg=mycolor)

frameForFirstPage=Frame(window2,bg=mycolor)
frameForFirstPage.grid(row=1,column=1)

def signup():                                                               # for signup
    bookpyxl=openpyxl.load_workbook("RentalDatabase.xlsx")
    mysheet=bookpyxl["Users"]
    mycolor="grey14"
    labelColor="AntiqueWhite3"



    wbook=xlrd.open_workbook("RentalDatabase.xlsx")
    xlsheet=wbook.sheet_by_name("Users")

    frameForFirstPage.grid_forget()

    frameForSignup=Frame(window2,bg=mycolor)
    frameForSignup.grid(row=1,column=1)


    window2.title("Sign up")

    def back():
        window2.title("CaRent!")
        
        frameForFirstPage.grid(row=1,column=1)
        frameForSignup.grid_forget()
    

    backButton=Button(frameForSignup,text="Back",command=back,width=7,bg="grey").grid(row=1,column=1,padx=5,pady=5,sticky="W")
    
    fname=Label(frameForSignup,text="First name:",bg=mycolor,fg=labelColor).grid(row=2,column=1,padx=5,pady=5)
    
    lname=Label(frameForSignup,text="Last name:",bg=mycolor,fg=labelColor).grid(row=3,column=1,padx=5,pady=5)
    
    username=Label(frameForSignup,text="User name:",bg=mycolor,fg=labelColor).grid(row=4,column=1,padx=5,pady=5)
    
    number=Label(frameForSignup,text="Phone number:",bg=mycolor,fg=labelColor).grid(row=5,column=1,padx=5,pady=5)
    
    email=Label(frameForSignup,text="Email:",bg=mycolor,fg=labelColor).grid(row=6,column=1,padx=5,pady=5)
    
    blank=Label(frameForSignup,bg=mycolor,fg=labelColor).grid(row=7,column=1,padx=5,pady=5)
    
    password=Label(frameForSignup,text="Password:",bg=mycolor,fg=labelColor).grid(row=8,column=1,padx=5,pady=5)
    
    blank=Label(frameForSignup,bg=mycolor,fg=labelColor).grid(row=9,column=1,padx=5,pady=5)
    
    rquestion=Label(frameForSignup,text="Recovery question:",bg=mycolor,fg=labelColor).grid(row=10,column=1,padx=5,pady=5)
    ranswer=Label(frameForSignup,text="Recovery answer:",bg=mycolor,fg=labelColor).grid(row=11,column=1,padx=5,pady=5)


    fname_e=Entry(frameForSignup,bg=mycolor,fg=labelColor)
    fname_e.grid(row=2,column=2,padx=5,pady=5)

    lname_e=Entry(frameForSignup,bg=mycolor,fg=labelColor)
    lname_e.grid(row=3,column=2,padx=5,pady=5)

    username_e=Entry(frameForSignup,bg=mycolor,fg=labelColor)
    username_e.grid(row=4,column=2,padx=5,pady=5)

    number_e=Entry(frameForSignup,bg=mycolor,fg=labelColor)
    number_e.grid(row=5,column=2,padx=5,pady=5)

    email_e=Entry(frameForSignup,bg=mycolor,fg=labelColor)
    email_e.grid(row=6,column=2,padx=5,pady=5)

    blank=Label(frameForSignup,bg=mycolor).grid(row=7,column=2,padx=5,pady=5)
    password_e=Entry(frameForSignup,bg=mycolor,fg=labelColor)
    password_e.grid(row=8,column=2,padx=5,pady=5)

    blank=Label(frameForSignup,bg=mycolor).grid(row=9,column=2,padx=5,pady=5)
    rquestion_e=Entry(frameForSignup,bg=mycolor,fg=labelColor)
    rquestion_e.grid(row=10,column=2,padx=5,pady=5)

    ranswer_e=Entry(frameForSignup,bg=mycolor,fg=labelColor)
    ranswer_e.grid(row=11,column=2,padx=5,pady=5)

    
    def phonecheck(x):
        listnum=[1,2,3,4,5,6,7,8,9,0]
        count=0
        for i in listnum:
            for j in x:
                if str(i)==str(j):
                    count+=1
          
        if count==len(x) and len(x)==10:
            register()
        else:
            messagebox.showerror("Error","Phone number cant contain letters and must be 10 digits.")
            




    def register():                          #xlsheet=Users
        nextrow=xlsheet.nrows+1
        a="A"+str(nextrow)   
        b="B"+str(nextrow)
        c="C"+str(nextrow)
        d="D"+str(nextrow)
        e="E"+str(nextrow)
        f="F"+str(nextrow)
        g="G"+str(nextrow)
        h="H"+str(nextrow)
        

        mysheet[a]=fname_e.get()
        mysheet[b]=lname_e.get()
        mysheet[c]=username_e.get()
        mysheet[d]=number_e.get()
        mysheet[e]=email_e.get()
        mysheet[f]=password_e.get()
        mysheet[g]=rquestion_e.get()
        mysheet[h]=ranswer_e.get()

        bookpyxl.save("RentalDatabase.xlsx") 
        messagebox.showinfo("Done","You have successfully signed up. ")

    def check():
        if fname_e.get()=="" or lname_e.get()=="" or username_e.get()=="" or number_e.get()=="" or email_e.get()=="" or password_e.get()=="" or rquestion_e.get()=="" or ranswer_e.get()=="":
            messagebox.showwarning("WARNING","All boxes must be filled.")

        elif (username_e.get()=="admin" or email_e.get()=="admin") and password_e.get()=="iamtheadmin":
            messagebox.showwarning("WARNING","Choose a different user name.")
        elif (username_e.get()=="dealer" or email_e.get()=="dealer") and password_e.get()=="iamthedealer":
            messagebox.showwarning("WARNING","Choose a different user name.")
             
        else:
            count=0
            for i in range(1,xlsheet.nrows):   #xlsheet=Users
                if username_e.get()==xlsheet.cell_value(i,2):
                    messagebox.showwarning("WARNING","Choose a different user name.")
                    count+=1
                    break
            if count==0:
                phonecheck(number_e.get())
            

    okay=Button(frameForSignup,text="Sign Up",width=40,height=2,bg="grey",command=check).grid(row=12,columnspan=4)

  





def login():                # for login
    wbook=xlrd.open_workbook("RentalDatabase.xlsx")
    xlsheet=wbook.sheet_by_name("Users")             #xlsheet=users
    count=1
    if uoremail_e.get()=="" or passrd_e.get()=="":
        messagebox.showwarning("Warning","You must fill out all the boxes!")
    elif uoremail_e.get()=="admin" or passrd_e.get()=="iamtheadmin":
        editcar()
    elif uoremail_e.get()=="dealer" or passrd_e.get()=="iamthedealer":
        pickcar()
        
    else:
        for i in range(1,xlsheet.nrows):
            if (uoremail_e.get()==xlsheet.cell_value(i,2) or uoremail_e.get()==xlsheet.cell_value(i,4)) and passrd_e.get()!=xlsheet.cell_value(i,5):
                messagebox.showwarning("Warning","Check username and password")
            elif (uoremail_e.get()==xlsheet.cell_value(i,2) or uoremail_e.get()==xlsheet.cell_value(i,4)) and passrd_e.get()==xlsheet.cell_value(i,5):
                thelist=["","","","","",'']
                thelist[0]=xlsheet.cell_value(i,0) #first name
                thelist[1]=xlsheet.cell_value(i,1) #last name
                thelist[2]=xlsheet.cell_value(i,3) #phone number
                thelist[3]=xlsheet.cell_value(i,4) #email
                thelist[4]=xlsheet.cell_value(i,5) #password
                thelist[5]="F"+str(i+1)            # location on user sheet for changing password
                user(thelist,window2,frameForFirstPage)
            
            elif uoremail_e.get()!=xlsheet.cell_value(i,2):
                count+=1
        if xlsheet.nrows==count:
            messagebox.showwarning("Warning","User doesn't exist.")





def forgotpass():                 #for forgetpass
    wbook=xlrd.open_workbook("RentalDatabase.xlsx")
    xlsheet=wbook.sheet_by_name("Users")
    for i in range(1,xlsheet.nrows):
        if uoremail_e.get()=="":
            messagebox.showwarning("Warning","Enter email or username.")
            break

        elif uoremail_e.get()==xlsheet.cell_value(i,2) or uoremail_e.get()==xlsheet.cell_value(i,4):
           window3=Tk()
           window3.resizable(0,0)
           mycolor="grey14"
           labelColor="AntiqueWhite3"
           window3.config(bg=mycolor)
           window3.title("Recovery Question")

           frameForQusetion=Frame(window3,bg=mycolor)
           frameForQusetion.grid(row=1,column=1) 
           
           intro=Label(frameForQusetion,text="Answer the security question",fg=labelColor,font="14",bg=mycolor).grid(row=1,column=1,columnspan=4,padx=7,pady=7)
           
           lab=Label(frameForQusetion,text=xlsheet.cell_value(i,6),font="14",bg=mycolor,fg=labelColor).grid(row=2,column=1,columnspan=4,padx=5,pady=5)
           
           anslabel=Label(frameForQusetion,text="The answer is",bg=mycolor,fg=labelColor).grid(row=3,column=1,padx=5,pady=5)
           
           ans=Entry(frameForQusetion,bg=mycolor,fg=labelColor)
           ans.grid(row=3,column=2,padx=5,pady=5)
           
           def checker():
                if ans.get()==xlsheet.cell_value(i,7):    # checking if revovery answer is the same
                
                    frameForQusetion.grid_forget() 
                    mycolor="grey14"
                    labelColor="AntiqueWhite3"
                    window3.title("Password")
                    frameForAnswer=Frame(window3,bg=mycolor)
                    frameForAnswer.grid(row=1,column=1)

                    def back():
                        window3.title("Recovery Question")
                        frameForAnswer.grid_forget()
                        frameForQusetion.grid(row=1,column=1)
                       

                    backButton=Button(frameForAnswer,text="Back",command=back,bg="grey").grid(row=1,column=1,padx=8,pady=8,sticky="W") 

                    passis=Label(frameForAnswer,text="Your password is",font=("Arial","10","bold"),bg=mycolor,fg=labelColor).grid(row=2,column=1,padx=8,pady=8)
                    itis=Label(frameForAnswer,text=xlsheet.cell_value(i,5),font=("Arial","10","bold"),fg="green",bg=mycolor).grid(row=2,column=2,padx=5,pady=1)
                else:
                    messagebox.showwarning("Incorrect","Your answer is incorrect.")
    
        
                   
           enter=Button(frameForQusetion,text="Enter",command=checker,width=35,height=2,bg="grey").grid(row=4,column=1,columnspan=4)   #button on forgot password window
               

    









labb=Label(frameForFirstPage,text="CaRent!",font=("Helvetica","30"),bg=mycolor,fg=labelColor).grid(row=1,column=1,columnspan=8,padx=10,pady=10)

uoremail=Label(frameForFirstPage,text="Email/username:",font=("Helvetica","15"),bg=mycolor,fg=labelColor).grid(row=2,column=1,padx=2,pady=2,sticky="W")
uoremail_e=Entry(frameForFirstPage,font=("Helvetica","13"),width=30,bg="grey15",fg=labelColor)
uoremail_e.grid(row=3,column=1,padx=5,pady=2)


passrd=Label(frameForFirstPage,text="Password:",font=("Helvetica","15"),bg=mycolor,fg=labelColor).grid(row=4,column=1,padx=2,pady=2,sticky="W")
passrd_e=Entry(frameForFirstPage,font=("Helvetica","13"),show="*",width=30,bg="grey15",fg=labelColor)
passrd_e.grid(row=5,column=1,padx=5,pady=2)


logn=Button(frameForFirstPage,text="Login",width=35,height=2,command=login,bg="grey").grid(row=6,column=1,columnspan=2,pady=5,padx=2)

sup=Button(frameForFirstPage,text="Sign Up",width=20,height=2,command=signup,bg="grey").grid(row=7,column=1,pady=7,padx=5,sticky="W")

forgot=Button(frameForFirstPage,text="Forgot password",width=20,height=2,command=forgotpass,bg="grey").grid(row=7,column=1,pady=7,padx=5,sticky="E")



window2.mainloop()







