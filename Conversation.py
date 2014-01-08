'''
Conversation Object is an array of Message objects

@author: admin
'''
import datetime
from Message import *
import xlwt

now = datetime.datetime.now()
class Conversation:
    
    def __init__(self):
        self.messages = []
        
    def add_message(self, message):
        self.messages.append(message)
        
    def print_conversation(self):
        
        for msg in self.messages:
            print(msg.print_message())
        
    #Returns the row where it stops    
    def write_conversation_to_worksheet(self,row,worksheet):
        #print("COnvo is given: " + str(row))
        count = 0
        for msg in self.messages:
            if msg.empty:
                continue
            msg.write_message_to_worksheet(row + count, worksheet);
            count = count + 1
        
        return count + row
        