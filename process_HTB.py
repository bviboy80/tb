import os
import sys
import csv
import re
import time
import datetime
import shutil
import zipfile
import PyPDF2
import smtplib
from email.mime.text import MIMEText
import win32com.client as win32


# Import custom modules
from logedit import Editor
import email_notification as en 


def main():
    currentyear = datetime.date.today().strftime('%Y')
    currentmonth = datetime.date.today().strftime('%m')
    currentday = datetime.date.today().strftime('%d')
    currentdate = datetime.date.today().strftime('%Y%m%d')


    ''' SELECT TEST OR PRODUCTION '''
    
    # PRODUCTION
    folder_paths = (r'F:\nycdof', r'P:\Vanguard\HTB', False)
    
    # TESTING
    # folder_paths = (r'P:\Vanguard\HTB\scripts\TESTING', r'P:\Vanguard\HTB\scripts\TESTING', True)
    
    
    
    
    ftp_main, main_folder, email_test = folder_paths
    
    # FTP FOLDERS
    ftp_peops_folder = os.path.join(ftp_main, "PEOPS")
    ftp_done_folder = os.path.join(ftp_peops_folder, "Done")

    # ARCHIVE/ERROR FOLDERS
    dup_folder = os.path.join(main_folder, "errorfiles", "duplicate", currentdate)
    error_folder = os.path.join(main_folder, "errorfiles", "bad", currentdate)

    # CURRENT JOB FOLDERS
    job_folder = os.path.join(main_folder, "PRINT_FILES", currentyear, currentdate)
    job_zipped_folder = os.path.join(job_folder, "zipped")
    job_print_folder = os.path.join(job_folder, "PRINT")

    print_file_list = []
    error_file_list = []

    # Initialize the Log Editor
    log_folder = os.path.join(main_folder, "log")
    log_editor = Editor(log_folder, currentyear)

    # filter for HTB Files only
    htb_files = searchFolderForHTBfiles(ftp_peops_folder)
    
    # Check for the good/correct files to process. Move bad/incorrect files to the archive
    # Each entry contains:
    # zip name, ftp path + filename, date received/modified, date processed/current date
    
    files_to_process = checkHTBfilesForErrors(htb_files, ftp_peops_folder,
                                              dup_folder, error_folder,
                                              error_file_list, log_editor, currentdate)

    # Process HTB files
    if len(files_to_process) > 0:
        if not os.path.exists(job_folder):
            os.makedirs(job_folder)
        if not os.path.exists(job_zipped_folder):
            os.mkdir(job_zipped_folder)
        if not os.path.exists(job_print_folder):
            os.mkdir(job_print_folder)
            
        for filename, ftp_file, date_rcvd, date_processed in files_to_process:
            
            # Move file to job and done folders
            job_zip_file = os.path.join(job_zipped_folder, filename)
            shutil.copy2(ftp_file, job_zip_file)
            shutil.move(ftp_file, os.path.join(ftp_done_folder, filename))

            # Unzip pdf file to print folder
            with zipfile.ZipFile(job_zip_file, 'r') as z:
                for unzipped_file in z.namelist():
        
                    print "\r\nUnzipping {}".format(filename)
                    z.extract(unzipped_file, job_print_folder)
                    
                    # Verify file is a pdf, then process 
                    test_file = os.path.join(job_print_folder, unzipped_file) 
                    print_file = os.path.join(job_print_folder, unzipped_file)
                    unzipped_pdf = None
                    
                    if fileIsPDF(test_file):
                        
                        # Add ".pdf" extension if the file does not have it 
                        unzipped_pdf = unzipped_file
                        if unzipped_file[-3:].upper() != "PDF" and unzipped_file[-1] != ".":
                            unzipped_pdf = unzipped_file + ".pdf"
                            print_file = os.path.join(job_print_folder, unzipped_pdf)
                            os.rename(test_file, print_file) 
                        
                        # Get PDF record counts
                        print "\r\nGetting record counts" 
                        with open(print_file, 'rb') as pf:
                            pdf_handle = PyPDF2.PdfFileReader(pf)
                            page_count = pdf_handle.getNumPages()
                            
                            # Get file info for email message and add to list
                            print_file_list.append((filename, unzipped_pdf, page_count, date_rcvd))
                            
                            # Add log entry
                            log_editor.addEntry([filename, unzipped_pdf, page_count, date_rcvd, date_processed, "PRINT"])
                    
                            print "\r\n{} sucessfully unzipped and processed".format(filename)
                            
                    # Process BAD PDFs
                    else:
                        z.close()
                        print "\r\nFile not a PDF or incorrect file type"
                        
                        # Delete bad PDF
                        os.remove(test_file)
                        
                        # Move BAD ZIP to error folder
                        if not os .path.exists(error_folder): 
                            os.mkdir(error_folder)
                        shutil.move(job_zip_file, os.path.join(error_folder, filename))
                        
                        # Add log entry
                        error_file_list.append([filename, date_rcvd, "BAD PDF/NOT A PDF"])
                        log_editor.addEntry([filename, unzipped_pdf, "", date_rcvd, date_processed, "BAD PDF/NOT A PDF"])
                        
                        
    if len(print_file_list) <= 0:
        email_test = True
    
    
    # send email with date, name of file, and counts to Vanguard
    print "Sending counts email to Vanguard"
    vg_email = en.VanguardEmail(print_file_list, currentdate, email_test)
    en.sendEmailMsgOL(vg_email.mailingList,
                 vg_email.msgSubject(),
                 vg_email.msgText())
    
    
    # Send email with date, file path, name of file, and counts to production
    print "Sending print files email to Production"
    prod_email = en.ProductionEmail(print_file_list, job_print_folder, currentdate, email_test)
    en.sendEmailMsgPY(prod_email.sender,
                 prod_email.mailingList,
                 prod_email.msgSubject(),
                 prod_email.msgText())
    
    
    # send email of error files to programmer
    print "Sending error email to Programmer"
    err_email = en.ErrorEmail(error_file_list, currentdate)
    en.sendEmailMsgPY(err_email.sender,
                 err_email.mailingList,
                 err_email.msgSubject(),
                 err_email.msgText())    

    
    
### END MAIN FUNCTION


                 

    
def searchFolderForHTBfiles(ftp_peops_folder):
    """ Search for HTB Only files by checking the FTP files/folders by name 
    Example file name: FTXPRODN.PROD.BF850W01.HTB-20180509-010016.zip """
    
    ftp_file_list = os.listdir(ftp_peops_folder)
    htb_files = []
    htb_file_pattern = re.compile(r'^FTXPRODN\.PROD\.BF850W01\.HTB-\d{8}-\d{6}\.zip($|.+)')
    
    for f in ftp_file_list:
        if htb_file_pattern.search(f) != None:
            htb_files.append(f)
        else:    
            print "\r\n{} not processed".format(f)
            continue

    return htb_files        


    
def checkHTBfilesForErrors(htb_files, ftp_peops_folder, dup_folder, error_folder, error_file_list, log_editor, currentdate):            
    """ Check each element in folder contents.
    If file does not follow the correct naming convention, skip it.
    If file is already in log, write duplicate in log and copy file to duplicate folder.
    If file is not a zip file, write incorrect file in log and copy file to error folder.
    If file is bad zip file, write bad zip file in log and copy file to error folder.
    All else, add to "files to process" list. """
    
    files_to_process = []
    
    for f in htb_files:
        
        ftp_file = os.path.join(ftp_peops_folder, f)
        dupfile = os.path.join(dup_folder, f)
        errfile = os.path.join(error_folder, f)
        
        # Format Date file received 
        date_mod = os.path.getmtime(ftp_file)
        dtup = datetime.datetime.fromtimestamp(date_mod).timetuple()
        f_date_mod_fmt = "{}{:0>2}{:0>2}".format(dtup.tm_year, dtup.tm_mon, dtup.tm_mday)

        # Check if file is a duplicate. Move to dup folder
        if log_editor.getEntry(f) != None:
            if not os.path.exists(dup_folder):
                os.makedirs(dup_folder)
            
            # Move incorrect file type to duplicate folder            
            shutil.move(ftp_file, dupfile)
            
            # Add duplicate to the log
            error_file_list.append([f, f_date_mod_fmt, "DUPLICATE"])
            log_editor.addEntry([f, "", "", f_date_mod_fmt , currentdate, "DUPLICATE"])    
            print "\r\n{} found in log. Copied to DUPLICATE folder".format(f)
        
        # Check if file is a not a zip file
        elif not zipfile.is_zipfile(ftp_file): 
            if not os.path.exists(error_folder):
                os.makedirs(error_folder)
            
            # Move incorrect file type to error folder
            shutil.move(ftp_file, errfile)
            
            # Add error file to the log
            error_file_list.append([f, f_date_mod_fmt, "INCORRECT FILE TYPE"])
            log_editor.addEntry([f, "", "", f_date_mod_fmt , currentdate, "INCORRECT FILE TYPE"])
            print "\r\n{} is incorrect file type. Copied to ERROR folder".format(f)
        
        else:
            zip_test = ""
            
            # Add to process list if all checks are passed   
            try:
                zip_test = zipfile.ZipFile(ftp_file)
                zip_test.close()
                files_to_process.append((f, ftp_file, f_date_mod_fmt, currentdate))
                print "\r\n{} successfully tested. Copied to Production folder".format(f)
            
            # Test if Zip file is bad. Move bad zip to error folder
            except zipfile.BadZipfile:
                zip_test.close()
                
                # Move bad zip file to error folder
                shutil.move(ftp_file, errfile)
                
                # Add bad zip to the log
                error_file_list.append([f, f_date_mod_fmt, "BAD ZIP"])
                log_editor.addEntry([f, "", "", f_date_mod_fmt, currentdate, "BAD ZIP"])
                print "\r\n{} is a BAD ZIP FILE. Copied to ERROR folder".format(f) 
                
    return files_to_process    


def fileIsPDF(test_file):            
    """ Use PDF Reader methods to test/verify file is a pdf """ 
    test_open = open(test_file, 'rb')
    try:
        open_pdf = PyPDF2.PdfFileReader(test_open)
        open_pdf.getNumPages()
        test_open.close()
        return True
    except:
        test_open.close()
        return False        
            
            
            
        
            
if __name__ == '__main__':
    main()
