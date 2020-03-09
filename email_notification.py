'''
Module that takes a list of HTB print files and formats and sends emails :

1) to Print Production with the files, records counts and print file location 
2) to Vanguard with the files and records counts
3) to the programmer with the files that had issues (i.e., bad zip file, bad PDF or no errors)


If sending email during testing, main program will set the "mail_test" variable to True.
Emails will only be send to the programmer
'''




import os
import sys
import csv
import re
import datetime
import shutil
import zipfile
import PyPDF2
import smtplib
from email.mime.text import MIMEText
import win32com.client as win32




class ProductionEmail(object):
    """ Create sender, recipient list, subject 
    and formatted message body for print production.  """
    
    def __init__(self, print_file_list, job_print_folder, currentdate, mail_test):
        self.print_file_list = print_file_list
        self.job_print_folder = job_print_folder
        self.currentdate = currentdate
        self.sender = "dsproduction@astfinancial.com"        
        self.mailingList = self.createMailingList(mail_test)       
    
    def createMailingList(self, mail_test):
        """ 'To' List"""
        
        mailList_dict = {
            True : ["sthomas@astfinancial.com"],
                         
            False : ["kmcneil@astfinancial.com","dschwarz@astfinancial.com",
                    "lszewc@astfinancial.com","sthomas@astfinancial.com",
                    "svogt@astfinancial.com","jherrera@astfinancial.com",
                    "norellana@astfinancial.com","lprice@astfinancial.com",
                    "rchetram@astfinancial.com","yjainarain@astfinancial.com"]
            }
        return mailList_dict[mail_test]
        
    def msgSubject(self):
        """ Email Subject Line """
        
        return "DOF HTB Bi-Weekly Mailing - PRINT FILES - {}".format(self.currentdate)
       
    def msgText(self):
        """Combine and format Mailing date and list of files with counts 
        into single message. Convert message to string."""
        
        # Set default if nothing to print
        messageText = "NOTHING TO PRINT."
        
        # Process if there are files to print    
        if len(self.print_file_list) > 0:
            
            # Add spaces in between file path for better readability 
            folder = " \\ ".join(self.job_print_folder.split("\\"))
            
            # Create string of print file list and print instructions
            filelist = ["{:<25}\t{:>7}\t{:>10}\t{}".format("PDF", "Records", "Date Recvd", "From FTP Zip File")]
            totalcounts = 0
            
            for zipf, pdf, counts, recvd in self.print_file_list:
                totalcounts += counts
                filelist.append("{:<25}\t{:>7}\t\t{:>10}\t{}\r\n".format(pdf, counts, recvd, zipf))
            
            filelist.append("\r\n\r\nTotal Records: {}".format(totalcounts))
            filelist.append("\r\n\r\nPrint simplex on 20 lb., 8.5\" x 11\", white paper")
            filelist.append("\r\n1 impression (1 sheet) per record")
            filelistStr = "\r\n".join(filelist)

            # Combine folder path and print list into single string
            messageText = "File(s) below are available at :\r\n" + \
                "\r\n" + \
                folder + "\r\n" + \
                "\r\n" + \
                "\r\n" + \
                filelistStr
                
        return messageText     


        
class VanguardEmail(object):
    """ Create sender, recipient list, subject 
    and formatted message body for Vanguard.  """
    
    def __init__(self, print_file_list, currentdate, mail_test):
        self.print_file_list = print_file_list
        self.currentdate = currentdate
        self.sender = "sthomas@astfinancial.com"        
        self.mailingList = self.createMailingList(mail_test)       
    
    def createMailingList(self, mail_test):
        """ 'To' List"""
        
        mailList_dict = {
            True : ["sthomas@astfinancial.com"],
                         
            False : ["sthomas@astfinancial.com","abhagwandin@vanguarddirect.com",
            "mmuniz@vanguarddirect.com","kswan@hellovanguard.com","Team3@hellovanguard.com"]
            }
        return mailList_dict[mail_test]
        
    def msgSubject(self):
        """ Email Subject Line """

        return "DOF HTB Bi-Weekly Mailing - COUNTS - {}".format(self.currentdate)
       
    def msgText(self):
        """Combine and format Mailing date and list of files with counts 
        into single message. Convert message to string."""
        
        # Set default if nothing to print
        messageText = "NOTHING TO PRINT."
        
        # Process if there are files to print    
        if len(self.print_file_list) > 0:
            
            # Create string of print file list and print instructions
            filelist = ["{:<25}\t{:>7}\t\t\t{:>10}\t{}".format("PDF", "Records", "Date Recvd", "From FTP Zip File")]
            totalcounts = 0
            
            for zipf, pdf, counts, recvd in self.print_file_list:
                totalcounts += counts
                filelist.append("{:<25}\t{:>7}\t\t{:>10}\t{}\r\n".format(pdf, counts, recvd, zipf))
            
            filelist.append("\r\n\r\nTotal Records: {}".format(totalcounts))
            filelistStr = "\r\n".join(filelist)

            # Combine folder path and print list into single string
            messageText = "File(s) below were received for processing:\r\n" + \
                "\r\n" + \
                "\r\n" + \
                filelistStr
                
        return messageText        


class ErrorEmail(object):
    """ Create sender, recipient list, subject 
    and formatted message body for production.  """
    
    def __init__(self, error_file_list, currentdate):
        self.error_file_list = error_file_list
        self.currentdate = currentdate
        self.sender = "dsproduction@astfinancial.com"        
        self.mailingList = self.createMailingList()       
    
    def createMailingList(self):
        """ 'To' List"""

        return ["sthomas@astfinancial.com"]
        
    def msgSubject(self):
        """ Email Subject Line """

        return "DOF HTB Bi-Weekly Mailing - ERRORS - {}".format(self.currentdate)
       
    def msgText(self):
        """Combine and format Mailing date and list of files with counts 
        into single message. Convert message to string."""
        
        # Set default if nothing to print
        messageText = "NO ERRORS"
        
        # Process if there are files to print    
        if len(self.error_file_list) > 0:
            
            # Create string of print file list and print instructions
            filelist = ["{:<95}{:>5}{:>10}{:>10}{}".format("Zip File","", "Date Recvd","", "Error")]
            
            for zipf, recvd, err in self.error_file_list:
                filelist.append("{:<60}{:>5}{:>10}{:>10}{}\r\n".format(zipf, "", recvd, "", err))
            filelistStr = "\r\n".join(filelist)

            # Combine folder path and print list into single string
            messageText = "File(s) below produced the following errors :\r\n" + \
                "\r\n" + \
                "\r\n" + \
                filelistStr
                
        return messageText

        
def sendEmailMsgOL(mailingList, msgSubject, msgText):
    """ Send emails. Takes in objects that already contain:
    sender, recipient list, subject and message body. 
    Uses Outlook to send to non-AST emails. """
    
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = ";".join(mailingList)
    mail.Subject = msgSubject
    mail.Body = msgText
    mail.Send()


def sendEmailMsgPY(sender, mailingList, msgSubject, msgText):
    """ Send emails. Takes in objects that already contain:
        sender, recipient list, subject and message body. """
    
    msg = MIMEText(msgText)
    msg['From'] = sender
    msg['To'] = ",".join(mailingList)
    msg['Subject'] = msgSubject

    server = ""
    recipients_not_mailed = ""
    
    try:
        # Send the email via AST SMTP server.
        server = smtplib.SMTP("Mail3.amstock.com")
        recipients_not_mailed = server.sendmail(sender, mailingList, msg.as_string())
        
        if len(recipients_not_mailed) == 0:
            print recipients_not_mailed
            print "All recipients were successfully contacted"
        else:
            print "Message sent. The following recipients were rejected:"
            for addresses, errorcode in recipients_not_mailed.items():
                print "{}: {}".format(addresses, errorcode)
        
        server.quit()
        
    except smtplib.SMTPSenderRefused as sr:
        print sr
        server.quit()
    
    except smtplib.SMTPRecipientsRefused as rr:
        print rr
        print "Following recipients were rejected:"
        for addresses, errorcode in recipients_not_mailed.items():
            print "{}: {}".format(addresses, errorcode)
        
        server.quit()
    
    except smtplib.SMTPDataError as de:
        print de
        server.quit()
        
    except smtplib.SMTPException as se:
        print se
        server.quit()
        
    except:
        print "Unknown error. Unable to send email"
        server.quit()
        
        
