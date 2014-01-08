'''
Message object is a tuple of

sender
date sent
time sent
message body

@author: admin
'''
import datetime

now = datetime.datetime.now()
class Message:
    
    SENDER =0
    RECIPIENT = 1
    DATE = 2
    TIME = 3
    MSG = 4
    
    def __init__(self, sender, recipient, date,time,msg):
        
        #Excel sheet row constant
        
        self.sender = sender
        self.recipient = recipient
        self.date=date
        self.time = time
        self.msg = msg
        
        if(len(self.sender) == 0):
            self.empty = True
        else:
            self.empty = False
        

    def print_message(self):
        return "From: " + self.sender +" To: " + self.recipient + " " + self.date + " " + self.time + " " + self.msg
    
    '''
    writes this message into worksheet at the prescribed row
    '''
    def write_message_to_worksheet(self, row, worksheet):
        if len(self.sender) == 0:
            return
        print("Row: " + str(row))
        worksheet.write(row, self.SENDER, self.sender )
        worksheet.write(row, self.RECIPIENT, self.recipient )
        worksheet.write(row, self.DATE, self.date )
        worksheet.write(row, self.TIME, self.time )
        worksheet.write(row, self.MSG, self.msg )

