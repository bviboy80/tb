import os
import sys
import csv
import smtplib
from email.mime.text import MIMEText
import collections


""" Module to open, read and write to log file(s)."""

class Editor(object):
    
    def __init__(self, log_folder, currentyear):
        self.log_folder = log_folder    
        self.currentyear = currentyear
        self.current_log = self.getCurrentLog()
        
        
    def getCurrentLog(self):
        """ Create new log for the year if none exist  """
        
        curr_log = os.path.join(self.log_folder, "HTB_{}.log".format(self.currentyear))
        
        if not os.path.exists(curr_log):
            with open(curr_log, 'wb') as w:
                logWriter = csv.writer(w, delimiter="|")
                header = ["ZIP Filename", "PDF",
                          "Record Counts", "Date Received",
                          "Date Processed", "Status"] 
                logWriter.writerow(header)
        return curr_log
       
       
    def getAllEntries(self):
        """ Read log files and create a dict of zip files processed """
        
        log_entries_dict = collections.defaultdict(list)
        for logfile in os.listdir(self.log_folder):
            log = os.path.join(self.log_folder, logfile)
            with open(log, 'rb') as l:
                logCSVreader = csv.reader(l, delimiter="|")
                logCSVreader.next() # skip header
                try:
                    for row in logCSVreader:
                        zip_file = row[0]
                        log_entries_dict[zip_file].append(row)
                except:
                    pass
        return log_entries_dict            
    
    
    def getEntry(self, zip_file):
        """ Get log entry for a file"""
        
        return self.getAllEntries().get(zip_file)

      
    def addEntry(self, entry):
        """ Append entry to the log """
        
        with open(self.current_log, 'ab') as a:
            logAppender = csv.writer(a, delimiter="|")
            logAppender.writerow(entry)
         
