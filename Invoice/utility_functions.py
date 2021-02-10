# -*- coding: utf-8 -*-
"""
Created on 1st July 2020
@author: vinay.shetty@covalenseglobal.com
"""
import os
import sys
import re
import glob
import socket
import shutil
import logging
import win32gui
import time
import smtplib
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from functools import reduce
from pathlib import Path
import json
import subprocess as sp
import psutil

import pandas as pd
import sharepy
import pyautogui
#import pywinauto
import win32com.client as win32
from win32con import (SW_SHOW, SW_RESTORE, SW_MINIMIZE, SW_MAXIMIZE)
from win32com.client import Dispatch
from O365 import Account
from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient
import ast

class CommonFunctions():
    """
    A class containing common utility functions.
    ...
    Attributes:
    ----------
        audit_log (logging.Logger): Logger to catch audit logs/steps.
        error_log (logging.Logger): Logger to catch error logs.
    """

    def __init__(self, audit_log, error_log):
        '''
        The constructor for CommonFunctions class.

        Parameters:
            audit_log (logging.Logger): Logger to catch audit logs/steps.
            error_log (logging.Logger): Logger to catch error logs.
        '''

        self.audit_log = audit_log
        self.error_log = error_log
        
    def share_point_download(self, sp_username, sp_password, sp_link, flpath):
        '''
        The function is to download file from sharepoint.

            Parameters:
                sp_username(str): Username for logging in sharepoint.
                sp_password(str): Password to logging in sharepoint.
                sp_link(str): Url of sharepoint alongwith filename.
                flpath(str): Local path alongwith filename wherer file is to be downloaded.

            Returns:
                Boolean: True if succesfully downloaded file else False.
        '''

        for download_cnt in range(5):
            try:
                sess = sharepy.connect(site='unilever.sharepoint.com',
                                    username=sp_username, password=sp_password)
                if sess.getfile(sp_link, filename=flpath).status_code == 200:
                    sess.getfile(sp_link, filename=flpath)
                    sess.close()
                    return True
                else:
                    sess.close()
            except:
                pass
        self.error_log.error(f'Error while downloading {sp_link}', exc_info=True)
        return False

    def share_point_upload(self, sp_username, sp_password, file_path, path_initial, sp_folder, file_name, folder_create=True):
        '''
        The function is to upload file in sharepoint.

            Parameters:
                sp_username(str): Username for logging in sharepoint.
                sp_password(str): Password to logging in sharepoint.
                file_path(str): Path of file to upload.
                path_initial(str): Initial path of sharepoint after .com.
                sp_folder(str): Remaining path of folders after path_initial.
                file_name(str): Name of the file to be uploaded as in sharepoint.
                folder_create(Boolean): Will create folder in sharepoint as per sp_folder if True. 
                                        By default set as True.

            Returns: Boolean (True/False)
        '''
        for upload_cnt in range(5):
            try:
                sess = sharepy.connect(site='unilever.sharepoint.com', username=sp_username, password=sp_password)
                headers = {"accept": "application/json;odata=verbose","content-type": "application/x-www-urlencoded; charset=UTF-8"}

                if folder_create:   # Checking if folder needs to be created in sharepoint
                    self.audit_log.info('Creating folder in sharepoint.')
                    url_start = 'https://unilever.sharepoint.com'
                    folder_split = sp_folder.split('/')
                    full_url = url_start + '/' + '/'.join(folder_split[:])
                    if not sess.get(full_url).status_code == 200:   # Checking if full path is not present
                        for folder_loop in range(len(folder_split)):   # Looping through total folders
                            if folder_loop:
                                full_url = url_start + '/' + '/'.join(folder_split[:-folder_loop])
                                if sess.get(full_url).status_code == 200:   break   # Exiting if folder present

                        while True: # Looping till all folders are created
                            if folder_loop:
                                folder_path = '/'.join(folder_split[2: 1 - folder_loop])
                            else:
                                folder_path = '/'.join(folder_split[2: ])
                            sess.post(url_start + '/' + '/'.join(folder_split[0:2]) + "/_api/web/folders",
                            json={
                                "__metadata": {"type": "SP.Folder"},
                                "ServerRelativeUrl": folder_path
                            })
                            if not folder_loop: break   # Exiting after last folder creation
                            folder_loop -= 1

                with open(file_path, 'rb') as read_file:
                    content = read_file.read()
                status = sess.post("https://unilever.sharepoint.com/"+ path_initial + "/_api/web/GetFolderByServerRelativeUrl('/"+sp_folder+"')/Files/add(url='"+file_name+"',overwrite=true)", data=content, headers = headers)
                self.audit_log.info(f'Status Code - {str(status)} while uploading {file_name} to sharepoint {sp_folder}')
                if status.status_code is 200:
                    return True
            except:
                pass

        self.error_log.error(f'Error while uploading file - {file_name} to the sharepoint {sp_folder}', exc_info=True)
        return False

    def share_point_upload_new(self, sp_username, sp_password, local_file_path, upload_url, folder_create=True):
        '''
        The function is to upload file in sharepoint.

            Parameters:
                sp_username(str): Username for logging in sharepoint.
                sp_password(str): Password to logging in sharepoint.
                local_file_path(str): Path of file to upload alongwith filename.
                upload_url(str): Complete sharepoint url path along with filename.
                folder_create(Boolean): Will create folder in sharepoint as per sp_folder if True. 
                                        By default set as True.

            Returns: Boolean (True/False)
        '''
        url_split = upload_url.split("/")
        path_initial = '/'.join((url_split[3:5]))
        sp_folder = '/'.join(url_split[3:-1])
        file_name = url_split[-1]
        return self.share_point_upload(sp_username, sp_password, local_file_path, path_initial, sp_folder, file_name, folder_create=folder_create)

    def create_folder(self, path_list, del_folder=False):
        '''
        The function is to create folder.

            Parameters:
                path_list(list): A list of folder path to be created. 
                del_folder(Boolean): Will delete folder if exists, if set True.
                                     By default set as False.

            Returns: None
        '''

        try:
            for path in path_list:
                if os.path.exists(path) and del_folder:
                    shutil.rmtree(path)
                Path(path).mkdir(parents=True, exist_ok=True)
        except:
            self.error_log.error(
                f'Error while creating folder {path}', exc_info=True)

    def open_refresh_save_xl(self, file_list):
        '''
        The function is to open excel file, refresh all, save and close.

            Parameters:
                file_list(list): A list of excel file path to be opened.

            Returns:
                Boolean: True if successfully opened excel file, refeshed, saved and closed else False.
        '''

        # Checking if files are present
        got_error = False
        for files in file_list:
            if not os.path.isfile(files):
                self.audit_log.info(f'{os.path.basename(files)} is missing.')
                got_error = True
        if got_error:
            return False    # Exiting if workbook is missing

        try:
            for files in file_list:
                Application = Dispatch("Excel.Application")
                Application.Visible = False
                Application.DisplayAlerts = False
                Application.AskToUpdateLinks = False
                Application.EnableEvents = False
                self.audit_log.info(f'Opening file {os.path.basename(files)}')
                workbook = Application.Workbooks.Open(files)
                workbook.RefreshAll()
                workbook.Save()
                self.audit_log.info(f'Saved file {os.path.basename(files)}')
                workbook.Close(False)
                Application.Quit()
                return True
        except:
            self.error_log.error(
                f'Error while saving file {files}', exc_info=True)
            Application.Quit()
            return False

    def send_email(self, user, password, email_to, email_cc, subject, body, attachments=None, body_type='plain'):
        '''
        The function is to send email using O365.

            Parameters:
                user(str): Senders username.
                password(str): Senders password.
                email_to(str): Recipients email id for To.
                email_cc(str): Recipients email id for CC.
                subject(str): Email's subject.
                body(str): Email's body.
                attachments(list): List of files to be sent as attachment.
                body_typ(str): By default set as plain. Can be set as html if email body has tables.
            Returns: None.
        '''
        if socket.gethostname().lower().startswith('bnl'):
            try:
                self.audit_log.info('Sending email...')
                credentials = ('d39ac08f-0176-433b-a4d3-b47758d1d292', 'TVKT817nE_2m3nXo~-.44h-_1O~ePeKO9.')

                # the default protocol will be Microsoft Graph
                account = Account(credentials, auth_flow_type='credentials', tenant_id='f66fae02-5d36-495b-bfe0-78a6ff9f8e6e')
                if account.authenticate():
                    self.audit_log.info(f'connection to email server successful')
                else:
                    self.error_log.error(f'Error while connecting to email server', exc_info=True)
                
                server = account.new_message(resource=user) #os.environ['username'] + "@unilever.com"
                # Assign email TO
                for each_to in re.split(',|;', email_to):
                    server.to.add(each_to)
                # Assign email TO
                for each_cc in re.split(',|;', email_cc):
                    server.cc.add(each_cc)

                server.subject = subject+": Dev"
                body = body+'\n'+socket.gethostname()

                attachment_path = os.path.abspath(os.path.dirname(__file__))+"\\UnileverSignature.png"
                server.attachments.add(attachment_path)
                att = server.attachments[0]
                att.is_inline = True
                att.content_id = 'image.png'
                attachment = '<br><img src="cid:%s">' % att.content_id
                body = body + '<br>' + attachment
                server.body = body
                # add email attachment
                for file_attach in attachments or []:
                    server.attachments.add(file_attach)
                    # m.attachments.add(f'C:\\Users\\CVLUSER\\Desktop\\python\\Test.xlsx')
                server.send()
                self.audit_log.info('Sent email successfully')

            except:
                self.error_log.error(f'Error while sending email', exc_info=True)
        else:
            try:
                self.audit_log.info('Sending email...')
                server = smtplib.SMTP('smtp-in.unilever.com: 587')
                server.starttls()
                server.login(user, password)

                message = MIMEMultipart()
                message['From'] = user
                message['To'] = email_to
                message['Cc'] = email_cc
                message['Subject'] = subject

                message.attach(MIMEText(body, body_type))

                for files in attachments or []:
                    try:
                        with open(files, "rb") as fil:
                            part = MIMEApplication(
                                fil.read(),
                                Name=os.path.basename(files)
                            )
                        # After the file is closed
                        part['Content-Disposition'] = 'attachment; filename="%s"' % os.path.basename(files)
                        message.attach(part)
                    except:
                        self.error_log.error(f'Error while attaching {files} to email', exc_info=True)
                
                server.sendmail(user, re.split(',|;', email_to)+re.split(',|;', email_cc), message.as_string())
                self.audit_log.info('Sent email successfully')
            except:
                self.error_log.error(f'Error while sending email', exc_info=True)
            finally:
                server.quit()

    def send_email_with_image(self, user, password, email_to, email_cc, subject, body, image,attachments=None, body_type='plain'):
        '''
        The function is to send email using O365 with inline image.

            Parameters:
                user(str): Senders username.
                password(str): Senders password.
                email_to(str): Recipients email id for To.
                email_cc(str): Recipients email id for CC.
                subject(str): Email's subject.
                body(str): Email's body.
                image: file path with file name
                attachments(list): List of files to be sent as attachment.
                body_typ(str): By default set as plain. Can be set as html if email body has tables.
            Returns: None.
        '''
        try:
            self.audit_log.info('Sending email...')
            credentials = ('d39ac08f-0176-433b-a4d3-b47758d1d292', 'TVKT817nE_2m3nXo~-.44h-_1O~ePeKO9.')

            # the default protocol will be Microsoft Graph
            account = Account(credentials, auth_flow_type='credentials', tenant_id='f66fae02-5d36-495b-bfe0-78a6ff9f8e6e')
            if account.authenticate():
                self.audit_log.info(f'connection to email server successful')
            else:
                self.error_log.error(f'Error while connecting to email server', exc_info=True)
            
            server = account.new_message(resource=user) #os.environ['username'] + "@unilever.com"
            # Assign email TO
            for each_to in re.split(',|;', email_to):
                server.to.add(each_to)
            # Assign email TO
            for each_cc in re.split(',|;', email_cc):
                server.cc.add(each_cc)

            server.subject = subject+": Dev"

            
            server.attachments.add(image)
            att1 = server.attachments[0]
            att1.is_inline = True
            att1.content_id = 'image.png'
            attachment = '<br><img src="cid:%s">' % att1.content_id

            if body.find('Regards') is not -1:
                pos = body.find('Regards')
                body = body[:pos] + '<br>' + attachment + '<br>' + body[pos:]
            else:
                body = body + '<br><br>' + attachment

            body = body+'\n'+socket.gethostname()

            attachment_path = os.path.abspath(os.path.dirname(__file__))+"\\UnileverSignature.png"
            server.attachments.add(attachment_path)
            att2 = server.attachments[1]
            att2.is_inline = True
            att2.content_id = 'sigature.png'
            attachment = '<br><img src="cid:%s">' % att2.content_id
            body = body + '<br>' + attachment
            server.body = body
            # add email attachment
            for file_attach in attachments or []:
                server.attachments.add(file_attach)
                # m.attachments.add(f'C:\\Users\\CVLUSER\\Desktop\\python\\Test.xlsx')
            server.send()
            self.audit_log.info('Sent email successfully')

        except:
            self.error_log.error(f'Error while sending email', exc_info=True)

    def email_data(self, file, sheet, row):
        '''
        The function is to read email data from mapping file.

            Parameters:
                file(str): Excel file path to be read.
                sheet(str): Sheet name of the file to be read.
                row(str): Row no. in the sheet to be read.
            Returns:
                Tuple: Data as Email To, CC, Subject, Body
        '''

        try:
            content = pd.read_excel(file, sheet_name=sheet)
            content = content.fillna('')    # To avoid nan error
            to = content.iat[row, 1]
            cc = content.iat[row, 2]
            subject = content.iat[row, 3]
            body = content.iat[row, 4]
            return to, cc, subject, body
        except:
            self.error_log.error(
                f'Error while reading email data from {os.path.basename(file)}, Sheet: {sheet}, row no. {row+2}', exc_info=True)

    def kill_process(self, process_list):
        '''
        The function is to terminate the process.

            Parameters:
                process_list(list): List of processes to be terminated.
            Returns: None
        '''
        for each_proc in psutil.process_iter():
            try:
                if each_proc.name() in process_list:
                    each_proc.kill()
            except:
                self.error_log.error(
                    f'Error while killing process', exc_info=True)

    def copy_paste(self, copy_wb, copy_sheet, copy_range, paste_wb, paste_sheet, paste_range, wait_sec=3600):
        '''
        The function copies data from one excel file to another.

            Parameters:
                copy_wb(str): Excel file path from which data has to be copied.
                copy_sheet(str): Name of sheet in the file to be copied.
                copy_range(str): Range in excel file from where data has to be copied.
                paste_wb(str): Excel file path where data has to be copied.
                paste_sheet(str): Name of sheet in the file to be pasted.
                paste_range(str): Range in excel file where data has to be pasted.
                wait_sec(int): Maximum time to wait for copy. Set to 1 hour by default.
            Returns:
                Boolean: True if data copy paste success else False.
        '''

        try:
            if not wait_sec:
                wait_sec = 3600  # If waiting time for copy paste is 0 then setting 60 mins
            # Function to convert column alphabet into number
            def col2num(col): return reduce(lambda x, y: x*26 + y,
                                            [ord(c.upper()) - ord('A') + 1 for c in col])

            # Checking if files are present
            got_error = False
            for book in [copy_wb, paste_wb]:
                if not os.path.isfile(book):
                    self.audit_log.info(
                        f'{os.path.basename(book)} is missing.')
                    got_error = True
            if got_error:
                return False    # Exiting if workbook is missing

            Application = Dispatch("Excel.Application")
            Application.Visible = True
            Application.DisplayAlerts = False
            Application.AskToUpdateLinks = False
            Application.EnableEvents = False

            self.audit_log.info(
                f'Opening excel file for copying data - {os.path.basename(copy_wb)}')
            source = Application.Workbooks.Open(copy_wb)

            # Checking if sheet is present
            sheet_found = False
            for sheet_loop in range(source.Sheets.Count):
                if source.Sheets(sheet_loop+1).Name == copy_sheet:
                    sheet_found = True
                    break

            if not sheet_found:
                self.audit_log.info(
                    f'{copy_sheet} sheet is missing in {os.path.basename(copy_wb)}.')
                Application.Quit()
                return False

            copy_sht_ref = source.Worksheets(copy_sheet)
            Application.Worksheets(copy_sheet).Activate()

            try:    # Ask Forgiveness approach
                # Will fail if row no. is not present in copy range at end
                int(copy_range[-1])
            except:
                # If colon is present Only taking data after colon and removing character at end
                colm = copy_range.split(':')[-1][:-1]
                colm_no = col2num(colm)
                last_row = copy_sht_ref.Cells(
                    copy_sht_ref.Rows.Count, colm_no).End(-4162).Row
                copy_range = copy_range[:-1]+str(last_row)
                time.sleep(4)

            start_time = time.time()
            copy_sht_ref.Range(copy_range).Copy()

            # Calculating total rows to be copied
            try:
                if len(copy_range.split(':')) > 1:
                    copy_row = int(re.findall(r'\d+', copy_range.split(':')[1])[0]) - int(
                        re.findall(r'\d+', copy_range.split(':')[0])[0])
                else:
                    copy_row = 1
            except:
                copy_row = 0

            #   Work around for waiting till Application.StatusBar property is read
            while True:
                try:
                    while Application.StatusBar in ['Processing...', 'Busy']:
                        break
                    else:
                        break
                except:
                    pass

                # Exiting if time taken to copy data exceeds time treshold
                end_time = time.time()
                if end_time - start_time > wait_sec:
                    self.audit_log.info(
                        f'Time taken to copy the data exceeded the time limit of {wait_sec} seconds - {os.path.basename(copy_wb)}')
                    Application.Quit()
                    return False

            # Exiting if time taken to copy data exceeds time treshold
            while Application.StatusBar in ['Processing...', 'Busy']:
                end_time = time.time()
                if end_time - start_time > wait_sec:
                    self.audit_log.info(
                        f'Time taken to copy the data exceeded the time limit of {wait_sec} seconds - {os.path.basename(copy_wb)}')
                    Application.Quit()
                    return False

            # To exclude time taken for operations like opening excel, validating workbook, sheets and finding last row
            pause_time_start = time.time()
            self.audit_log.info(
                f'Opening excel file for pasting data - {os.path.basename(paste_wb)}')
            dest = Application.Workbooks.Open(paste_wb)

            # Checking if sheet is present
            sheet_found = False
            for sheet_loop in range(dest.Sheets.Count):
                if dest.Sheets(sheet_loop+1).Name == paste_sheet:
                    sheet_found = True
                    break

            if not sheet_found:
                self.audit_log.info(
                    f'{paste_sheet} sheet is missing in {os.path.basename(paste_wb)}.')
                Application.Quit()
                return False

            paste_sht_ref = dest.Worksheets(paste_sheet)
            Application.Worksheets(paste_sheet).Activate()

            # If colon is present then taking data only before colon or complete range
            colm = paste_range.split(':')[0]
            try:    # Ask Forgiveness approach
                # Will fail if row no. is not present in copy range
                int(colm[-1])
                # Extracting row no. from range
                row_no = re.findall(r'\d+', colm)[0]
                # Extracting and converting column into no. from range
                colm_no = col2num(colm.replace(row_no, ''))
            except:   # Calculating last row no
                # column refernce in no. after removing last character
                colm_no = col2num(colm[:-1])
                last_row = paste_sht_ref.Cells(
                    paste_sht_ref.Rows.Count, colm_no).End(-4162).Row
                row_no = last_row + 1   # Incrementing to paste from blank cell after last row

            try:
                total_row = copy_row + int(row_no)
            except:
                total_row = 0

            if total_row > 1048576:
                self.audit_log.info(
                    f'CRITICAL: Data is exceeding total no. of rows 1048576 while copying from {os.path.basename(copy_wb)}, Sheet: {copy_sheet} to {os.path.basename(paste_wb)}, Sheet: {paste_sheet}')
                return False

            # To exclude time taken for operations like opening excel, validating workbook, sheets and finding last row
            pause_time_end = time.time()
            # Resetting start counter by adding pause time
            start_time = start_time + pause_time_end - pause_time_start
            paste_sht_ref.Cells(row_no, colm_no).Select()
            paste_sht_ref.Paste()

            #   Work around for waiting till Application.StatusBar property is read
            while True:
                try:
                    while Application.StatusBar in ['Processing...', 'Busy']:
                        break
                    else:
                        break
                except:
                    pass

                # Exiting if time taken to paste data exceeds time treshold
                end_time = time.time()
                if end_time - start_time > wait_sec:
                    self.audit_log.info(
                        f'Time taken to paste the data exceeded the time limit of {wait_sec} seconds - {os.path.basename(paste_wb)}')
                    Application.Quit()
                    return False

            # Exiting if time taken to paste data exceeds time treshold
            while Application.StatusBar in ['Processing...', 'Busy']:
                end_time = time.time()
                if end_time - start_time > wait_sec:
                    self.audit_log.info(
                        f'Time taken to paste the data exceeded the time limit of {wait_sec} seconds - {os.path.basename(paste_wb)}')
                    Application.Quit()
                    return False

            source.Close(False)
            dest.Close(True)
            Application.Quit()
            self.audit_log.info('Copy Paste data success.')
            return True
        except:
            self.error_log.error(
                f'Pasting data failed for {os.path.basename(paste_wb)}', exc_info=True)
            Application.Quit()
            return False

    def bex_refresh(self, env, client, lang, user, password, bex_analyzer_path, hkey_1, hkey_2, sys_no, bex_data, tot_tab, query_file, output_path, refresh_time_sec, chk_sheet, chk_row, chk_col, chk_error):
        '''
        The function opens bex analyzer and then excel query file. Then using add-ins change variable values.

            Parameters:
                env(str): SAP Environment(eg. .B2P [B2P_Unity_Team]).
                client(str): SAP Client (eg. 110).
                lang(str): SAP Logon Language.
                user(str): Username generated from Token.
                password(str): Password generated from Token.
                bex_analyzer_path(str): Path of BEx Analyzer Excel(.XLA) file.
                htkey_1(str): Hot key value to be used while opening add-ins Change variable.
                htkey_2(str): Hot key value to be used while opening add-ins Change variable.
                sys_no(str): BEx System no. used for connection.
                bex_data(list): List of data to be pasted in form after opening change variable from add-in.
                tot_tab(int): Total no of manual tabs needed to reach "Check" button after opening change variable from add-in.
                query_file(str): Excel file complete path in which report is to be populated.
                output_path(str): Excel file complete path where excel file is to be saved after populating data.
                refresh_time_sec(int): Time in seconds to wait for BEx Analyzer to generate report.
                chk_sheet(str): Sheet in query_file where BEx Analyzer will populate data.
                chk_row(int): Row no. to check for if any errors after running BEx analyzer.
                chk_col(int): Column no. to check for if any errors after running BEx analyzer.
                chk_error(list): List of errors to check if BEx refresh failed.
            Returns:
                str/Boolean: 'Login failed' - Unable to establish BEx connection using credentials.
                             'Refresh failed' - Unable to populate report/data after change variables in BEx analyzer.
                             True: If successfully populated data/report after change variables in BEx analyzer.
                             False: If failed to populate data/report after change variables in BEx analyzer.

        '''
        try:
            self.kill_process(['excel.exe', 'EXCEL.EXE', 'saplogon.exe'])
            if not refresh_time_sec:
                refresh_time_sec = 3600
            self.audit_log.info('Starting BEx Refresh module....')
            Application = win32.Dispatch("Excel.Application")
            Application.Visible = True
            Application.DisplayAlerts = False
            Application.AskToUpdateLinks = False
            Application.EnableEvents = False
            Application.IgnoreRemoteRequests = True

            # Now that Excel is open, open the BEX Analyzer Addin xla file
            Application.Workbooks.Open(bex_analyzer_path)

            # Run the SetStart macro that comes with BEX so it pays attention to you
            Application.Run("BExAnalyzer.xla!SetStart")

            # Logon directly to BW using the sapBEXgetConnection macro
            myConnection = Application.Run(
                "BExAnalyzer.xla!sapBEXgetConnection")

            myConnection.Client = str(client)
            myConnection.user = user
            myConnection.password = password
            myConnection.language = lang
            myConnection.systemnumber = sys_no
            myConnection.system = env
            myConnection.SAProuter = ""
            myConnection.Logon(0, 1)

            if myConnection.IsConnected:
                self.audit_log.info('BEX Connection : SUCCESS')
            else:
                self.error_log.error('BEX Connection : FAILED', exc_info=True)
                return 'Login failed'

            # Now initialize the connection to make it actually usable
            self.audit_log.info('Initializing Connection')
            Application.Run("BExAnalyzer.xla!sapBEXinitConnection")

            # Now open the file you want to refresh
            self.audit_log.info(
                f'Loading the input file in BEX from loc {query_file}')
            bex_query = Application.Workbooks.Open(query_file, 0, False)
            sht_ref = bex_query.Worksheets(chk_sheet)
            Application.Worksheets(chk_sheet).Activate()
            Application.Run("BExAnalyzer.xla!SAPBEXrefresh", True)

            exlFilename = os.path.basename(query_file) + ' - Excel'
            self.audit_log.info('Bringing Excel file to the front.')
            windowID = pywinauto.findwindows.find_window(title=exlFilename)
            time.sleep(5)
            win32gui.ShowWindow(windowID, SW_RESTORE)
            time.sleep(1)
            win32gui.ShowWindow(windowID, SW_SHOW)
            time.sleep(1)
            win32gui.ShowWindow(windowID, SW_MINIMIZE)
            time.sleep(1)
            win32gui.ShowWindow(windowID, SW_MAXIMIZE)
            time.sleep(5)

            try:
                self.audit_log.info('Updating bex parameters process start.')
                win32gui.SetForegroundWindow(windowID)
                time.sleep(5)
                self.audit_log.info('Trying to open "Change Variable" menu.')
                pyautogui.FAILSAFE = False
                pyautogui.hotkey('alt', 'x')
                time.sleep(2)
                pyautogui.press(hkey_1)
                time.sleep(2)
                pyautogui.press(hkey_2)
                time.sleep(20)

                # Changing variable values from excel bex analysis toolbar add-ins
                self.audit_log.info(
                    'Trying to enter the values in bex analysis toolbar.')
                for loop in range(tot_tab):
                    pyautogui.press('tab')
                    if bex_data[loop]:
                        pyautogui.write(bex_data[loop])
                    time.sleep(2)

                # click on check and then ok
                pyautogui.press('enter')
                time.sleep(4)
                self.audit_log.info("Parameters are checked")
                for i in range(0, 2):
                    pyautogui.hotkey('shift', 'tab')
                    time.sleep(4)
                pyautogui.press('enter')
                time.sleep(4)
                self.audit_log.info("Parameters are ok")
            except Exception:
                self.error_log.error(
                    'Error occurred while checking the bex parameters', exc_info=True)
                return 'Refresh failed'

            start_time = time.time()
            time.sleep(20)

            time_exceed = False
            break_loop = True
            while True:  # Work around for waiting till Application.StatusBar property is read
                try:
                    while Application.StatusBar in ['Processing...', 'Busy']:
                        end_time = time.time()
                        if end_time - start_time > refresh_time_sec:  # Checking for time treshold
                            self.audit_log.info(
                                'Refresh failed as exceeded the time treshold.')
                            time_exceed = True
                            Application.Quit()
                            self.kill_process(['saplogon.exe'])
                    else:
                        print(Application.StatusBar)
                        break_loop = True   # For exiting while true loop
                except:
                    pass

                if time_exceed:  # Exiting if time treshold time exceeded
                    return 'Time exceeded'
                if break_loop:
                    break   # Exiting while true loop
            time.sleep(20)
            for check in chk_error:  # Checking for error after refresh
                if check in sht_ref.Cells(chk_row, chk_col).Value:
                    self.audit_log.info(
                        f'Refresh failed. {sht_ref.Cells(chk_row, chk_col).Value}')
                    Application.DisplayAlerts = False
                    Application.Quit()
                    self.kill_process(['saplogon.exe'])
                    return 'Data unavailable'

            try:
                Application.DisplayAlerts = False
                bex_query.SaveAs(output_path)
                bex_query.Close(False)
                Application.Quit()
                self.kill_process(['saplogon.exe'])
                return True
            except:
                self.error_log.error(
                    f'Unable to save {output_path}.', exc_info=True)
                Application.DisplayAlerts = False
                Application.Quit()
                self.kill_process(['saplogon.exe'])
                return False

        except Exception as err:
            self.error_log.error(f'CRITICAL: {err}', exc_info=True)
            Application.DisplayAlerts = False
            Application.Quit()
            self.kill_process(['saplogon.exe'])
            return False

    def bex_refresh_unattended(self, env, client, lang, user, password, bex_analyzer_path, sys_no, query_file, output_path, refresh_time_sec, chk_sheet, bex_param_range):
        '''
        The function opens bex analyzer and then excel query file. Then using add-ins change variable values.

            Parameters:
                env(str): SAP Environment(eg. .B2P [B2P_Unity_Team]).
                client(str): SAP Client (eg. 110).
                lang(str): SAP Logon Language.
                user(str): Username generated from Token.
                password(str): Password generated from Token.
                bex_analyzer_path(str): Path of BEx Analyzer Excel(.XLA) file.
                sys_no(str): BEx System no. used for connection.
                query_file(str): Excel file complete path in which report is to be populated.
                output_path(str): Excel file complete path where excel file is to be saved after populating data.
                refresh_time_sec(int): Time in seconds to wait for BEx Analyzer to generate report.
                chk_sheet(str): Sheet in query_file where BEx Analyzer will populate data.
                bex_param_range(excel range): Excel range with parameters to be passed to change variables window
            Returns:
                str/Boolean: 'Login failed' - Unable to establish BEx connection using credentials.
                             'Refresh failed' - Unable to populate report/data after change variables in BEx analyzer.
                             True: If successfully populated data/report after change variables in BEx analyzer.
                             False: If failed to populate data/report after change variables in BEx analyzer.

        '''
        try:
            if not refresh_time_sec:
                refresh_time_sec = 3600
            self.audit_log.info('Starting BEx Refresh module....')
            Application = win32.Dispatch("Excel.Application")
            Application.Visible = True
            Application.DisplayAlerts = False
            Application.AskToUpdateLinks = False
            Application.EnableEvents = False
            Application.IgnoreRemoteRequests = True

            # Now that Excel is open, open the BEX Analyzer Addin xla file
            Application.Workbooks.Open(bex_analyzer_path)

            # Run the SetStart macro that comes with BEX so it pays attention to you
            Application.Run("BExAnalyzer.xla!SetStart")
            time.sleep(10)

            # Logon directly to BW using the sapBEXgetConnection macro
            myConnection = Application.Run(
                "BExAnalyzer.xla!sapBEXgetConnection")
            time.sleep(10)
            # assign connection objects values
            myConnection.applicationserver = "130.24.5.116"
            try:
                myConnection.Client = str(client)
            except:
                myConnection.client = str(client)
            myConnection.user = user
            myConnection.password = password
            myConnection.language = lang
            myConnection.systemnumber = sys_no
            myConnection.system = env
            myConnection.SAProuter = ""
            myConnection.Logon(0, 1)
            time.sleep(10)

            if myConnection.IsConnected == 1:
                self.audit_log.info('BEX Connection : SUCCESS')
            else:
                self.error_log.error('BEX Connection : FAILED', exc_info=True)
                return 'Login failed'

            # Now initialize the connection to make it actually usable
            self.audit_log.info('Initializing Connection')
            Application.Run("BExAnalyzer.xla!sapBEXinitConnection")
            time.sleep(10)

            # Now open the file you want to refresh
            self.audit_log.info(
                f'Loading the input file in BEX from loc {query_file}')
            bex_query = Application.Workbooks.Open(query_file, 0, False)
            sht_ref = bex_query.Worksheets(chk_sheet)
            Application.Worksheets(chk_sheet).Activate()
            try:
                Application.Run(
                    "BExAnalyzer.xla!SAPBEXsetVariables", bex_param_range)
                time.sleep(10)
            except:
                self.audit_log.info(f'Refresh failed')
                Application.DisplayAlerts = False
                Application.Quit()
                self.kill_process(['saplogon.exe'])
                return 'Time exceeded'
            Application.Run("BExAnalyzer.xla!SAPBEXrefresh", True)
            time.sleep(10)
            try:
                Application.DisplayAlerts = False
                bex_query.SaveAs(output_path)
                bex_query.Close(False)
                Application.Quit()
                self.kill_process(['saplogon.exe'])
                return True
            except:
                self.error_log.error(
                    f'Unable to save {output_path}.', exc_info=True)
                Application.DisplayAlerts = False
                Application.Quit()
                self.kill_process(['saplogon.exe'])
                return False
        except Exception as err:
            self.error_log.error(f'CRITICAL: {err}', exc_info=True)
            Application.DisplayAlerts = False
            Application.Quit()
            self.kill_process(['saplogon.exe'])
            return False

    def login_details(self, login_file_path=None, token=None, sp_username=None, sp_password=None):
        '''
        The function helps getting login credentials for respective token id.
            Parameters:
                login_file_path(str): sharepont link where credentials file is saved.
                token(str): Token id to generate credentials.
                sp_username(str) : Sharepoint username
                sp_password(str): Sharepoint password.
            Returns:
                str: Success - It return login credentials as list
                     Failure - It return empty list [].
            Eg: details = login_details("IA-EU-C1P-01")
        '''
        try:
            self.audit_log.info(f'Read credentials from key vault server')
            if socket.gethostname().lower().startswith('bnl'):
                credential = DefaultAzureCredential()
                KVUri="https://bnlwe-da04-d-901334-kv02.vault.azure.net/"
                client = SecretClient(vault_url=KVUri, credential=credential)
                # token="Lev-Token2"
                retrieved_secret = client.get_secret(token)
                details = ast.literal_eval(retrieved_secret.value)
                try:
                    return [details['username'],details['Password1'],details['Password2']]
                except:
                    try:
                        return [details['username'],details['Password1']]
                    except:
                        return [details['Password1']]
            else:
                vault_name = "bnlwe-da04-d-901334-kv02"
                p = sp.Popen(["powershell.exe",
                r'az keyvault secret show --vault-name ' + vault_name + ' --name ' + token],
                stdout = sp.PIPE,
                stderr = sp.PIPE)
                result,err = p.communicate()
                details = json.loads(json.loads(result.decode('utf-8').replace("'", '"'))['value'])
                try:
                    return [details['username'],details['Password1'],details['Password2']]
                except:
                    try:
                        return [details['username'],details['Password1']]
                    except:
                        return [details['Password1']]
        except Exception as err:
            self.error_log.error(f'CRITICAL: {err}', exc_info=True)
            return []

    def get_sp_folderlist(self, link, username, password, site, library, relative_path):
        """
        Function to get a list of files and folders in a Sharepoint path.
        Arguments:
            link {str} -- Link of the Sharepoint Site.
            username {str} -- Username for Sharepoint Login
            password {str} -- Password for Sharepoint Login
            site {str} -- link + the site of the Sharepoint location
            library {str} -- Library of the files. Can get this from left pane on
                Sharepoint website.
            relative_path {str} -- Relative Path of the folder from where to
                get the files.

        Returns:
            list -- Returns a list of Files/Folder in Sharepoint Location.
        """
        session = sharepy.connect(
            site=link, username=username, password=password)

        # Count of JSON returns
        itemCount = 5000
        list1 = []
        condt = True
        id = ""
        while condt:
            link = ("{}/_api/web/lists/getbytitle('{}')/items"
                    "?$select=FileLeafRef,FileRef"
                    ",Id&$top={}&%24skiptoken=Paged%3DTRUE%26p_ID%3D{}"
                    ).format(site, library, itemCount, id)

            files = session.get(link).json()["d"]["results"]
            list1 = list1 + files
            id = files[-1]["Id"]

            if (len(files) != itemCount):
                condt = False

        output_list = []
        for file in list1:
            fullpth = file["FileRef"]
            if (fullpth.startswith(relative_path)) & (fullpth != relative_path):
                output_list.append(file["FileRef"])
        return output_list

    def share_point_get_folder(self, foldername, sp_username, sp_password, sp_path, input_path):
        """
        Function to get a list of files and folders in a Sharepoint path.
        Arguments:
            foldername -- sharepoint folder name
            sp_username -- Username for Sharepoint Login
            sp_password -- Password for Sharepoint Login
            sp_path     -- sharepoints folder path where files are placed
            input_path  -- Input local folder path where files to be downloaded
        Returns:
            Bool value -- True for Success/False for Failure.
        Example:
            bool_val = getByFolder('sp_foldername','iarp.dev5@unilever.com','Blripa@04',
                                    r"/SNCWebPortal/Shared Documents/Working/BW Files/", r'E:\test')
        """
        try:
            splink = "https://unilever.sharepoint.com"
            site = splink + "/sites/"+sp_path.split('/')[1]
            # Always use documents
            library = (r"Documents")
            # Sharepoint path
            relative_path = (r"/sites"+sp_path)

            result = self.get_sp_folderlist(link=splink, username=sp_username, password=sp_password,
                                            site=site, library=library,
                                            relative_path=relative_path)
            print("\n".join(result))
            print(input_path+'\\'+foldername)
            for file in result:
                url_split = file.split('/')
                file_name = url_split[len(url_split) - 1]
                self.share_point_download(sp_username, sp_password, splink + relative_path + '/' + file_name,
                                          os.path.join(input_path+'\\'+foldername, file_name))
            return True
        except Exception as err:
            self.error_log.error(f'CRITICAL: {err}', exc_info=True)
            return False

    def power_query_refresh(self, file_name):
        """
        Function to perform power query refresh.
        Arguments:
            file_name -- workbook file name where power query refresh to perform 
        Returns:
            Bool value -- True for Success/False for Failure.
        Example:
            bool_val = pwq_refresh_file('test.xlsm')
        """
        try:
            Application = win32.Dispatch('Excel.Application')
            Application.Visible = True
            Application.DisplayAlerts = False
            Application.AskToUpdateLinks = False
            Application.EnableEvents = False
            refresh_file_wb = Application.Workbooks.Open(file_name)
            for each_query in refresh_file_wb.Connections:
                bg_query_flag = each_query.OLEDBConnection.BackgroundQuery
                time.sleep(5)
                each_query.OLEDBConnection.BackgroundQuery = False
                time.sleep(5)
                each_query.Refresh()
                time.sleep(5)
                each_query.OLEDBConnection.BackgroundQuery = bg_query_flag
                time.sleep(5)
            refresh_file_wb.Save()
            time.sleep(5)
            refresh_file_wb.Close()
            Application.Quit()
            return True
        except Exception as err:
            self.error_log.error(f'CRITICAL: {err}', exc_info=True)
            return False
