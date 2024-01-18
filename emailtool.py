from tkinter import *
class outlook:
    def __init__(self):       
        self.excelFile = None
        self.senderEmail = None
        self.columnName = None
        self.subject = None
        self.body = None
        self.attachment = None
        self.checkboxID = None
        self.__password = None
    
    def send(self):
        import smtplib, ssl , email
        from email import encoders
        from email.mime.base import MIMEBase
        from email.mime.multipart import MIMEMultipart
        from email.mime.text import MIMEText
        import pandas as pd
        from tkinter import messagebox
        from socket import gaierror

        sender_email = self.senderEmail
        excel_file = self.excelFile
        colname = self.columnName
        subject = self.subject
        body = self.body
        attachment = self.attachment
        checkboxid = self.checkboxID

        
        df = pd.read_excel(excel_file)
        receiver_emails = list(df[colname])
        
        message = MIMEMultipart()
        message["From"] = sender_email
        message["To"] = ", ".join(receiver_emails)
        message["Subject"] = subject
       
        message.attach(MIMEText(body, "plain"))

        if len(attachment) != 0:
            name = attachment.split("/")[-1]
            with open(attachment, "rb") as attachment:
                # Add file as application/octet-stream
                # Email client can usually download this automatically as attachment
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())

            encoders.encode_base64(part)

            part.add_header(
                "Content-Disposition",
                f"attachment; filename= {name}",
            )

            message.attach(part)
        
        text = message.as_string()
        context = ssl.create_default_context()
        try:
            with smtplib.SMTP("smtp.office365.com", 587) as server:
                server.ehlo()
                server.starttls(context=context)
                server.ehlo()
                server.login(sender_email,self.__password)
                server.sendmail(sender_email,receiver_emails,text)
        except smtplib.SMTPAuthenticationError:
            messagebox.showerror(title='Warning',message='Please recheck email and password')
        except gaierror:
            messagebox.showerror(title='Warning',message='Please check your internet connect')



    def gui(self):
        from tkinter import ttk,filedialog
        from tkinter.filedialog import askopenfilename
        from tkinter import scrolledtext
        from tkinter import messagebox
        root = Tk() 

        root.title("Mailing Bot")

        root.geometry("1000x600+70+70")         #920x500+100+100
        f = Frame(root,width=700,height=600)

        senderemail = StringVar()
        filepath2 = StringVar()
        subject = StringVar()
        body = StringVar()
        filepath3 = StringVar()
        s = IntVar()

        def columnvalues():         #function for getting all columns names stored it into combobox
            import pandas as pd
            global excelFile
            excelFile = str(filepath2.get())
            df = pd.read_excel(str(filepath2.get()))
            list1 = list(df.columns)
            return list1

        def open_file():   # function for browse button to browse file
            fi2 = filedialog.askopenfilename(initialdir='/',title='Select cache path',filetypes=(('excel files','*.xlsx'),('all files','*.*')))
            filepath2.set(fi2)
            list1 = columnvalues()
            comb(list1)

        l1 = Label(f,text="Enter Sender email ID:",font=14)
        e1 = Entry(f,font=14,width=35,textvariable=senderemail)       
                
        l2 = Label(f,text='Enter your excel file path for receiver emails:',font=14)
        e2 = Entry(f,font=14,width=35,textvariable=filepath2)     # input for excel file

        style = ttk.Style()
        style.theme_use('alt')
        style.configure('TButton',background='#00b4d8')
        brow2 = Button(f,text='Browse',width=8,command=open_file,font=14,bg='#00b4d8',relief='raised')  #ttk.Button(f,text='Browse',command=open_file)   # browse button 

        l3 = Label(f,text='Select your receiver emails column :',font=14)
        # comobobox
        global n
        n = StringVar()
        combo1 = ttk.Combobox(f,font=14,textvariable=n) 
        combo1.grid(row=2,column=1,pady=20,sticky='W')
        combo1.config(state='disabled')
        def comb(list1):     # function for all columns name display in combobox after browse excel file
            # print(list1)
            combo1['values'] = tuple(list1)
            combo1.config(state='eabled')

        def value():   # function call when particular checkbutton is click
            def image():        # this function display gui when image checkbutton is click
                global l5
                global e3
                global brow3
                l5 = Label(f,text='Select path of your Image:',font=14)
                def open_image():
                    im = filedialog.askopenfilename(initialdir='/',title='select image file',filetypes=(('jpg','*.jpg'),('jpeg','*.jpeg'),('png','*.png'),('all image','*.*')))
                    filepath3.set(im)

                e3 = Entry(f,font=14,width=35,textvariable=filepath3)
                brow3 = Button(f,text='Browse',command=open_image,width=8,font=14,bg='#00b4d8',relief='raised')

                l5.grid(row=6,column=0,pady=20)
                e3.grid(row=6,column=1,pady=20,sticky='W')
                brow3.grid(row=6,column=1,columnspan=2,sticky='E')


            def document(): # this function display gui when document checkbutton is click
                global l5
                global e3
                global brow3
                l5 = Label(f,text='Select document path:',font=14)
                def open_document():
                    fi3 = filedialog.askopenfilename(initialdir='/',title='select document',filetypes=(('excel file','*.xlsx'),('pdf','*.pdf'),('doc file','*.docx'),('text file','*.txt'),('ppt','*.pptx'),('all file','*.*')))
                    filepath3.set(fi3)

                e3 = Entry(f,font=14,width=35,textvariable=filepath3)
                brow3 = Button(f,text='Browse',font=14,width=8,bg='#00b4d8',relief='raised',command=open_document)
                l5.grid(row=6,column=0,pady=20)
                e3.grid(row=6,column=1,pady=20,sticky='W')
                brow3.grid(row=6,column=1,columnspan=2,sticky='E')

            
            if s.get()==1:     # call image function when checkbutton value is 1
                image()  
                c2.config(state='disabled')
            elif s.get()==2:    # call document function when checkbutton value is 2
                document()
                c1.config(state='disabled')

        def cancel():   # this function reset the checkbutton and destory particular checkbutton gui
            if s.get()==1 or s.get()==3:   
                s.set(0)
                l5.destroy()
                e3.destroy()
                brow3.destroy()
                c2.config(state='normal')
            elif s.get()==2 or s.get()==4:
                s.set(0)
                l5.destroy()
                e3.destroy()
                brow3.destroy()
                c1.config(state='normal')
        
        def confirm():              # this function for validation 
            self.columnName = str(n.get())
            self.senderEmail = str(senderemail.get())
            self.excelFile = str(filepath2.get())
            self.subject = str(subject.get())
            self.body = str(t1.get("1.0","end-1c"))
            self.attachment = str(filepath3.get())
            self.checkboxID = s.get()
            if len(senderemail.get()) == 0:
                messagebox.showerror(title='Warning',message='Please enter SenderEmail ID')
            elif len(filepath2.get()) == 0:
                messagebox.showerror(title='Warning',message='select porper attachment')
            elif len(n.get()) == 0:
                messagebox.showerror(title='Warning',message='select correct option')
            elif len(subject.get()) == 0:
                messagebox.showerror(title='Warning',message='Please enter Subject')
            elif len(t1.get("1.0","end-1c")) == 0:
                messagebox.showerror(title='Warning',message='Please enter Body')
            else:
                e1.config(state='disabled')
                e2.config(state='disabled')
                brow2.config(state='disabled')
                combo1.config(state='disabled')
                sb.config(state='disabled')
                t1.config(state='disabled')
                c1.config(state='disabled')
                c2.config(state='disabled')
                b1.config(state='disabled')
                if s.get()==1 or s.get()==3 or s.get()==2 or s.get()==4:
                    e3.config(state='disabled')
                    brow3.config(state='disabled')
                # password window will open 
                topl = Toplevel()
                topl.title("Mailing Bot")
                topl.geometry("600x150+300+200")
                f1 = Frame(topl,width=400,height=100)
                l = Label(f1,text="Enter sender email Password:",font=12)
                l.grid(row=0,column=0,pady=20)
                pd = StringVar()
                e4 = Entry(f1,font=14,width=35,textvariable=pd,show="*")
                e4.grid(row=0,column=1,pady=20)
                def submit():
                    if len(pd.get()) == 0:
                        messagebox.showerror(title='Warning',message='Please enter Password')
                    else:
                        self.__password = str(pd.get())
                        topl.destroy()

                btn = Button(f1,text="submit",font=14,width=8,justify=RIGHT,bg='#00b4d8',command=submit)
                btn.grid(row=2,column=1,sticky='W',pady=20)
                f1.pack()
                topl.mainloop() 
                # close window

            
        def reset():
            senderemail.set('')
            filepath2.set('')
            combo1.set('')
            subject.set('')
            s.set(0)
            combo1.config(state='disabled')
            filepath3.set('')
            e1.config(state='normal')
            e2.config(state='normal')
            brow2.config(state='normal')
            sb.config(state='normal')
            t1.config(state='normal')
            t1.delete("1.0","end")
            c1.config(state='normal')
            c2.config(state='normal')
            b1.config(state='normal')
            e3.config(state='normal')
            brow3.config(state='normal')
            l5.destroy()
            e3.destroy()
            brow3.destroy()

        sub = Label(f,text='Subject :',font=14)
        sb = Entry(f,font=14,width=35,textvariable=subject)
            
        l4 = Label(f,text="Body",font=14)
        t1 = scrolledtext.ScrolledText(f,font=14,width=50,height=5)  

        l6 = Label(f,text='Select what you want to send in outlook:',font=14,padx=10)
        c1 = Checkbutton(f,text='Image',variable=s,font=14,onvalue=1,offvalue=3,command=value)
        c2 = Checkbutton(f,text='Document',variable=s,font=14,onvalue=2,offvalue=4,command=value)
        b1 = Button(f,text='Cancel',width=8,bg='#00b4d8',font=14,command=cancel) 


        btn1 = Button(f,text='Confirm',font=14,width=8,justify=RIGHT,bg='#00b4d8',command=confirm)
        btn2 = Button(f,text='Reset',font=14,width=8,justify=RIGHT,bg='#00b4d8',command=reset)
        btn3 = Button(f,text='Send',font=14,width=8,justify=RIGHT,bg='#00b4d8',command=self.send)  # start send message button
        
        l1.grid(row=0,column=0,pady=20)
        e1.grid(row=0,column=1,pady=20,sticky='W')
        l2.grid(row=1,column=0,pady=20)
        e2.grid(row=1,column=1,pady=20,sticky='W')
        brow2.grid(row=1,column=1,columnspan=2,pady=20,sticky='E')

        l3.grid(row=2,column=0,pady=20)

        sub.grid(row=3,column=0,pady=20)
        sb.grid(row=3,column=1,pady=20,sticky='W')
        l4.grid(row=4,column=0,pady=20)
        t1.grid(row=4,column=1,pady=20)


        l6.grid(row=5,column=0,pady=20)
        c1.grid(row=5,column=1,pady=20,sticky='W')
        c2.grid(row=5,column=1,pady=20,columnspan=2,sticky='W',padx=100)
        b1.grid(row=5,column=1,columnspan=3,pady=20,sticky='W',padx=220)

        btn1.grid(row=7,column=0,pady=20,sticky='E')   # confirm
        btn2.grid(row=7,column=1,pady=20)                # reset
        btn3.grid(row=7,column=2,sticky='W',pady=20)    # send

        f.pack()            

        root.mainloop()
    

o = outlook()
o.gui()