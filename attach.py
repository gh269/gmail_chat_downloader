'''
Created on Mar 19, 2013

Emails come out as quoted-printable encoded

'''

from xml.dom import minidom
from random import randint
import email, getpass, imaplib, os, re,quopri, xlwt,time
from Conversation import *
from Message import *

import sys, string, calendar
import Tkinter, tkFileDialog, tkMessageBox
import time
from tkCalendar import *


year = time.localtime()[0]
month = time.localtime()[1]
day =time.localtime()[2]
tk = Tkinter 

fnta = ("Times", 12)
fnt = ("Times", 14)
fntc = ("Times", 18, 'bold')

strtitle = "Calendar"
strdays = "Su  Mo  Tu  We  Th  Fr  Sa"
dictmonths = {'1':'Jan','2':'Feb','3':'Mar','4':'Apr','5':'May',
'6':'Jun','7':'Jul','8':'Aug','9':'Sep','10':'Oct','11':'Nov',
'12':'Dec'}

strdate = (str(day) +  "-" + dictmonths[str(month)] + "-" + str(year))
myFormats = [ ('Excel 97', '*.xls') ]
class ChatLogger:
    
    def __init__(self):
        return
        
    #Determine the encoding of a IMAP Message
    '''
    As input take the raw email body
    '''
    def determine_encoding(self,mail):
        first_occurence = mail.find('Content-Transfer-Encoding:')
        encoding_type = ""
        pattern = '^[\w-]+$'
        #Valid email if it contains the content transfer encoding string
        if first_occurence > -1:
            #Encoding is delimited by " ", so advance your pointer and concatentate
            #your substring until you run into a space
            string_index = len("Content-Transfer-Encoding:") + first_occurence + 1
            
            while string_index < len(mail):
                if re.match(pattern,mail[string_index]):
                    encoding_type += mail[string_index]
                    string_index += 1
                else:
                    break    
                
        return encoding_type

    #take a 7 bit encoded message and create conversation object
    #each email corresponds to a unique conversation object
    '''
    Extract contiguous section of XML formatted code
    this lasts 
    Tags desired: 
    Constructs a new conversation out of a, 
    XML string
    
    All times are given in GMT
    '''
    def new_conversation(self,mail):
        xmldoc = minidom.parseString(mail)
       
        #There should only be a single conversation in this string
        conversations = xmldoc.getElementsByTagName('con:conversation')
        if len(conversations) != 1:
            return -1
        c = Conversation()
        messages = conversations[0].getElementsByTagName('cli:message')
        for m in messages:
            sender = m.attributes['from'].value
            recipient = m.attributes['to'].value
            
            #in case sender or recipient has crap appended to the end
            edu = sender.find('edu')
            com = sender.find('com')
            
            tld_occurence = edu if edu >= 0 else com
            sender = sender[0:tld_occurence + len('com')]
            
            edu = recipient.find('edu')
            com = recipient.find('com')
            
            tld_occurence = edu if edu >= 0 else com
            recipient = recipient[0:tld_occurence + len('com')]     
            
            
            msglist = m.getElementsByTagName('cli:body')
            if len(msglist) <= 0:
                   continue
            else:
                msglist = msglist[0].childNodes
            out = []
            for node in msglist:
                out.append(node.data)
            msg = ''.join(out)
    
            datetime = m.getElementsByTagName('time')[0].attributes['ms'].value
            datetime = int(datetime) / 1000
            dt = time.strftime('%Y-%m-%d %I:%M%p', time.localtime(datetime))
            date = dt[0:10]
            timet = dt[11:len(dt)]
            
            message_to_add = Message(sender, recipient, date, timet, msg)
            c.add_message(message_to_add)
            
        return c

    '''
    Fetches XML section by grabbing subsection
    from one <con:conversation> tag to other 
    input is raw mail
    '''
    def fetch_xml_body(self,mail):
        open_conversation = mail.find('<con:conversation xmlns')
        close_tag = '</con:conversation>'
        close_conversation = mail.find(close_tag)
        
        return mail[open_conversation:close_conversation + len(close_tag)]
    
    def login(self,user,pwd):
        self.user = user
        self.pwd = pwd
    
    def setup_imap(self,imap_url):
        self.mailbox = imaplib.IMAP4_SSL(imap_url)
        
    def setup_chat_connection(self, user, pwd, imap_url):
        self.login(user, pwd)
        self.setup_imap(imap_url)
        try:
            self.mailbox.login(user, pwd)
        except:
            return -1
        #self.mailboxes = self.mailbox.list()
        self.mailbox.select("[Gmail]/Chats", readonly=True)
        return 0
    
    '''
        Google possible search terms:
            Person
            Date
            Subject
    '''
    def search_string(self, contact, date_from, date_to):
        if len(contact) == 0 and len(date_from) == 0 and len(date_to) == 0:
            return "ALL"
        search_str = "("
        if len(contact) > 0:
            search_str += 'FROM \"' + contact + '\" '
        if len(date_from) > 0:
            search_str += 'SINCE \"' + date_from + '\" '
        if len(date_to) > 0:
            search_str += 'BEFORE \"' + date_to + '\"'
        return search_str + ")"
    
    def create_spreadsheet(self, searchstr,savedir):
        resp, items = self.mailbox.search(None, searchstr)
        items = items[0].split()
        count = 0
        
        wb = xlwt.Workbook()
        Page = 1
        ws = wb.add_sheet('Logs' + str(Page))
        
        ws.write(0, Message.SENDER, 'Sender')
        ws.write(0, Message.RECIPIENT, 'Recipient')
        ws.write(0, Message.DATE, 'Date')
        ws.write(0, Message.TIME, 'Time')
        ws.write(0, Message.MSG, 'Message')
        
        counter = 1
        for emailid in items:
            resp, data = self.mailbox.fetch(emailid, "(RFC822)")
            email_body = data[0][1] # getting the mail content
            c = self.new_conversation(self.fetch_xml_body(quopri.decodestring(email_body)))
            convos = c.write_conversation_to_worksheet(counter, ws) 
            counter = convos
            if(counter > 65533):
                counter = 1
                Page = Page + 1
                ws = wb.add_sheet('Logs' + str(Page))
                ws.write(0, Message.SENDER, 'Sender')
                ws.write(0, Message.RECIPIENT, 'Recipient')
                ws.write(0, Message.DATE, 'Date')
                ws.write(0, Message.TIME, 'Time')
                ws.write(0, Message.MSG, 'Message')
            
            count = count + 1
            if count == 10:
                time.sleep(randint(0,3))
                count = 0
        
        try:
            wb.save(savedir)  
        except:
            return -1
        return 0
        




#a.setup_chat_connection(user,pwd,"imap.gmail.com")
#strtern = a.search_string("Sally Shi", "5-Jan-2013", "7-Jan-2013")
#a.create_spreadsheet(strtern, "Logs.xls")

class clsMainFrame(tk.Frame):
    def __init__(self, master):
        self.parent = master
        tk.Frame.__init__(master)
        self.date_var = tk.StringVar()
        self.date_var.set(strdate)  
        self.date_until = tk.StringVar()
        self.date_until.set(strdate)     
        '''
        Chats From
        '''
        self.name = "Chat from who?"
        self.entry = tk.Entry(master, textvariable=self.name)
        self.entry.pack(side="top", pady=20)
        self.entry.delete(0)
        self.entry.insert(0, self.name)
        self.entry.bind("<Button-1>", self.clearEntry)

        '''
        Chats Since
        '''
        self.chats_since_label = tk.Label(master, text= "Chats Since", bg = "white")
        self.chats_since_label.pack(side="top")  
        
        self.chats_since = tk.Label(master, textvariable= self.date_var, bg = "white")
        self.chats_since.pack(side="top", pady=10)
        self.chats_since.bind("<Button-1>",self.fnCalendar)
        '''
        Chats Until
        '''
        self.chats_until_label = tk.Label(master, text= "Chats Until", bg = "white")
        self.chats_until_label.pack(side="top")  
             
        self.chats_until = tk.Label(master, textvariable= self.date_until, bg = "white")
        self.chats_until.pack(side="top",pady=10)
        self.chats_until.bind("<Button-1>",self.fnCalendarUntil)
        
        '''
        Gmail Account Credentials
        '''
        self.gmail_address = "Enter Gmail Address"
        self.gmail = tk.Entry(master, textvariable = self.gmail_address)
        self.gmail.pack(side = "top", pady=15)
        self.gmail.bind("<Button-1>", self.clearAddr)
        self.gmail.delete(0)
        self.gmail.insert(0, self.gmail_address)
        
        self.gmail_password = "Enter Gmail Password"
        self.passw= tk.Entry(master, textvariable = self.gmail_password, show="*")
        self.passw.pack(side = "top", pady=15)
        self.passw.bind("<Button-1>", self.clearPass)
        self.passw.delete(0)
        self.passw.insert(0, self.gmail_password)
        
        
#         self.gmail_imap = "imap.gmail.com"
#         self.imap= tk.Entry(master, textvariable = self.gmail_imap)
#         self.imap.pack(side = "top", pady=15)
#         self.imap.bind("<Button-1>", self.clearImap)
#         self.imap.delete(0)
#         self.imap.insert(0, self.gmail_imap)
        
        
        downloadBtn = tk.Button(master, text = 'Download Chats',
         command = self.saveXLSSheet)
        downloadBtn.pack(side = 'bottom')

    def saveXLSSheet(self):
        
        file_name = tkFileDialog.asksaveasfilename(parent=root,filetypes=myFormats ,title="Save the logs as...")
        file_name = file_name + ".xls"
        a = ChatLogger()
        user = self.gmail.get()
        passw = self.passw.get()
        
        imap = "imap.gmail.com"
        d = a.setup_chat_connection(user, passw, imap)

        if d == -1:
            self.login_error()
            return
        from_field = self.entry.get()
        from_date = self.chats_since.cget('text')
        to_date = self.chats_until.cget('text')

        search_term = a.search_string(from_field,from_date, to_date)
        d = a.create_spreadsheet(search_term, file_name)
        if d == -1:
            self.save_error()
            return
        self.done(file_name)
        
    def login_error(self):
        tkMessageBox.showinfo("Login Error", "Error Logging In, Check Gmail Credentials")
    def save_error(self):
        tkMessageBox.showinfo("Login Error", "Error saving the spreadsheet, maybe a file with the same name is open?")
    def done(self, directory):
        tkMessageBox.showinfo("Done", "Spreadsheet of Logs saved to " + directory)
    def clearAddr(self,event):
        self.gmail.delete(0, len(self.gmail.cget('text')))
    def clearImap(self,event):
        self.imap.delete(0, len(self.gmail.cget('text'))) 
    def clearPass(self,event):
        self.passw.delete(0, len(self.passw.cget('text')))
    def clearEntry(self,event):
        self.entry.delete(0, len(self.entry.get()))
    def fnCalendar(self,event):
        tkCalendar(self.parent, year, month, day, self.date_var )
    def fnCalendarUntil(self,event):
        tkCalendar(self.parent, year, month, day, self.date_until )
        
        
        
        
if __name__ == '__main__':
    
    root =tk.Tk()
    root.title ("Gmail Chat Download")
    Frm = tk.Frame(root)
    clsMainFrame(Frm)
    Frm.pack()
    root.mainloop()

        
    