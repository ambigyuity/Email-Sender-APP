from tkinter import *
from tkinter.filedialog import askopenfilename
from email.message import EmailMessage
import os
import smtplib
import pandas as pd



#hwprlejjwaszlogs
#install pandas and xlrd

class EmailSender():
    def __init__(self):
        print('The init method')
    def emaillogin(self):
    #Setting username and password
        os.environ['EMAIL_USER']= self.userentry.get()
        os.environ['EMAIL_PASS']= self.passentry.get()
        self.EMAIL_ADDRESS= os.environ.get('EMAIL_USER')
        self.EMAIL_PASSWORD= os.environ.get('EMAIL_PASS')

        
        newmsg= EmailMessage()
        newmsg['Subject']= 'This is a placeholder'
        newmsg['From']= self.EMAIL_ADDRESS
        newmsg['To']= self.emaillist
        newmsg.set_content('This is the body')
        

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as ES:
        # STMP_SSL encrypts and identifies the traffic automatically
           # The code below is for smtplib.SMTP('smtp.gmail.com, 587) 
           # ES.ehlo() #identifiera
           # ES.starttls() #encrypt the traffic
           # ES.ehlo() #re-identify as an encrypted message

            ES.login(self.EMAIL_ADDRESS, self.EMAIL_PASSWORD)
            ES.send_message(newmsg)    

        
    def tkinter1(self):
        self.filename=''
        self.list1=[]
        self.list2=[]
        self.list3=[]
        self.list4=[]
        def openfile():
        
            self.filename = askopenfilename()
            excelfile=pd.read_excel(self.filename, "Sheet1")
            self.emaillist=excelfile['email address'].tolist()
            self.details=excelfile['filter'].tolist()
            self.moredet=excelfile['details'].tolist()
            populate(self.bodyframe)
            print(self.emaillist)

        def login(entry1,entry2):
                self.emaillogin()
                print('hello')
        def destroy():
            for x in self.list1:
                x.destroy()
            for y in self.list2:
                y.destroy()

        def getinfo():

            for b in self.list4:
                b.destroy()

            self.list4=[]

            populate(self.bodyframe)
            
            for row in range(len(self.emaillist)):

                md=self.moredet[row]
                self.moredet1=Label(self.bodyframe, text= md, bg= '#BFEFFF', anchor= 'nw')
                self.moredet1.grid(row=row, column =3)
                self.list4.append(self.moredet1)
                
        #have to alter the contents of self.details and self.moredet 
            
            
                
            
        def populate(frame):
            destroy()
            self.list1=[]
            self.list2=[]

            for row in range(len(self.emaillist)):
                self.NumberLabel=Label(self.bodyframe, text="%s" % (row + 1), width=3, borderwidth="1", relief="solid",bg='#BFEFFF')
                self.NumberLabel.grid(row=row, column=0)
                email_address= self.emaillist[row]
                self.EmailLabel=Label(self.bodyframe, text=email_address, bg='#BFEFFF',anchor='nw')
                self.EmailLabel.grid(row=row, column=1)
                self.list1.append(self.NumberLabel)
                self.list2.append(self.EmailLabel)
                
                
    
        def retrieve_val(event):
            myValue= variable.get()
            update_resultsfilter(myValue)

        def update_resultsfilter(newfilter):

            destroy()

            for b in self.list4:
                b.destroy()


            try:
                if newfilter=="All":
                    excelfile=pd.read_excel(self.filename, "Sheet1")
                    excelfile.set_index("filter", inplace=True)
                    self.emaillist=excelfile['email address'].tolist()
                    self.moredet=excelfile['details'].tolist()

                
                else:
                    self.emaillist=[]
                    self.moredet=[]
                    try:
                        excelfile=pd.read_excel(self.filename, "Sheet1")
                        excelfile.set_index("filter", inplace=True)
                        newdata=excelfile.loc[int(newfilter),['email address']]
                        detailsdata=excelfile.loc[int(newfilter),['details']]
                        self.moredet=detailsdata['details'].tolist()
                        self.emaillist=newdata['email address'].tolist()
                        


                        for b in self.list4:
                            b.destroy()
                    except:
                        excelfile=pd.read_excel(self.filename, "Sheet1")
                        excelfile.set_index("filter", inplace=True)
                        newdata=excelfile.loc[int(newfilter),['email address']]
                        detailsdata=excelfile.loc[int(newfilter),['details']]
                        self.emaillist.append(newdata['email address'])
                        self.moredet.append(detailsdata['details'])

                        for b in self.list4:
                            b.destroy()
      
                populate(self.bodyframe)
            except:
                print("No such filter")
                
            
                        
            
                


        
              
        window=Tk()
        window.geometry("600x700")
        variable=StringVar(window)
        variable.set("Filter")

    #UserFrame
        self.canvas=Canvas(window, height= 600, width= 700,bg='white')
        self.canvas.place(relheight=1, relwidth=1)
        

        self.userframe=Frame(self.canvas, bg= '#BFEFFF')
        self.userframe.place(relx=0.05, rely= 0.05, relheight= 0.1, relwidth= 0.9)

        self.userlabel=Label(self.userframe,text='Email:',bg='#BFEFFF')
        self.userlabel.place(relx=0.001, rely= 0.35)
        self.passlabel=Label(self.userframe, text='APP PW:', bg= '#BFEFFF')
        self.passlabel.place(relx= 0.42, rely= 0.35)
        
        
        self.userentry=Entry(self.userframe)
        self.userentry.place(relx=0.1, rely=0.05, relheight= 0.85, relwidth= 0.3)

        self.passentry=Entry(self.userframe)
        self.passentry.place(relx=0.55, rely=0.05, relheight= 0.85, relwidth= 0.4)

        self.userbutton= Button(self.canvas, text='Send', command= lambda: login(self.userentry.get(), self.passentry.get()))
        self.userbutton.place(relx=0.425, rely= 0.9, relheight= 0.05, relwidth= 0.1)

        self.filebutton=Button(self.canvas,text='Recipient List', command= lambda: openfile())
        self.filebutton.place(relx= 0.06, rely= 0.175, relheight= 0.06)
    #Subject and Body Frame
        self.bodycanvas=Canvas(self.canvas, bg='#BFEFFF', height= int(self.canvas['height'])*0.4, width= int(self.canvas['width'])*0.9)
        self.bodyframe= Frame(self.bodycanvas, bg= '#BFEFFF', height= int(self.bodycanvas['height'])*1, width= int(self.bodycanvas['width'])*1)
        self.bodycanvas.place(relx=0.05, rely= 0.25, relheight= 0.4, relwidth= 0.9)
        self.bodyframe.place(relx= 0, rely= 0, relheight= 1, relwidth= 1)

        self.scrollbar=Scrollbar(self.bodycanvas, orient= 'vertical', command= self.bodycanvas.yview)
        self.bodycanvas.configure(yscrollcommand=self.scrollbar.set)
        self.scrollbar.place(relx=0.978,rely=0.013, relheight= 0.975, relwidth= 0.02)

        self.bodycanvas.create_window((4,4),window=self.bodyframe, anchor= 'nw')
        self.bodyframe.bind("<Configure>", lambda event, canvas= self.bodycanvas: onFrameConfigure(self.bodycanvas))

       # self.bodylabel= Label(self.bodyframe, bd=5, text= '', bg='white', anchor= 'nw', justify='left')
       # self.bodylabel.place(relx= 0.0001, rely= 0.0001, relheight= 0.999999, relwidth= 1)
        OPTIONS= ["All", '1','2','3']

            
        self.menu=OptionMenu(window, variable, *OPTIONS, command= retrieve_val)
        self.menu.place(rely= 0.185, relx= 0.25)

        self.detailsbutton=Button(self.canvas, text= 'Details', command= lambda:getinfo())
        self.detailsbutton.place( relx= 0.39, rely= 0.185, relheight= 0.041, relwidth= 0.1)
     

        print(self.canvas['width'])
        def onFrameConfigure(canvas):
            canvas.config(scrollregion=canvas.bbox("all"))
        
            
        
 
                              
        
        window.mainloop()


        
        








me= EmailSender()
me.tkinter1()
