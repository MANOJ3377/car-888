import os
import sys
import socket
import logging
import configparser
import pandas as pd
import utility_functions

from datetime import datetime
from win32com.client import Dispatch

script_path = os.path.abspath(os.path.dirname(__file__))

current_time = datetime.now()
start_time = datetime.now()
mapping_folder = os.path.join(script_path, 'Mapping')
temp_folder = os.path.join(script_path, 'temp')
log_fol = os.path.join(script_path, "Logs")


def setup_logger(logger_name, log_file, level=logging.INFO):
    logger = logging.getLogger(logger_name)
    formatter = logging.Formatter(
        socket.gethostname() + ' : ' + '%(asctime)s : %(levelname)s : [%(filename)s:%(lineno)d] : %(message)s')

    fileHandler = logging.FileHandler(log_file, mode='w')

    fileHandler.setFormatter(formatter)

    streamHandler = logging.StreamHandler(sys.stdout)
    streamHandler.setFormatter(formatter)

    logger.setLevel(level)
    logger.addHandler(fileHandler)
    logger.addHandler(streamHandler)
    return logger

aud_log_file = "Mapping_Fin_TTS on Invoice_Task1_Audit_" + str(current_time.strftime("%Y-%m-%d_%H-%M")) + ".log"
aud_log_file = os.path.join(log_fol, aud_log_file)

err_log_file = "Mapping_Fin_TTS on Invoice_Task1_Error_" + str(current_time.strftime("%Y-%m-%d_%H-%M")) + ".log"
err_log_file = os.path.join(log_fol, err_log_file)

if not os.path.exists(log_fol):
    os.makedirs(log_fol)

audit_logger = setup_logger('audit', aud_log_file, level=logging.INFO)
error_logger = setup_logger('error', err_log_file, level=logging.ERROR)

audit_logger.info('Process started.')

# calling utility function
common = utility_functions.CommonFunctions(audit_logger, error_logger)

# Step 1: Creating the folder structure
input_path = os.path.join(script_path, 'Input')
output_path = os.path.join(script_path, 'Output')
temp_path = os.path.join(script_path, 'temp')

try:
    audit_logger.info('Step 1 - Creating folder')
    common.create_folder([input_path, output_path, temp_path], del_folder=True)
    audit_logger.info("Step 1  - Folder structure created")
except:
    error_logger.error("Error while creating folder structure",exc_info=True)
    audit_logger.info("Task Failed")
    sys.exit()

# Create log Folder Structure
# try:
#     Log_Path = os.path.join(script_path, "Logs")
#     if not os.path.exists(Log_Path):
#         os.makedirs(Log_Path, exist_ok=True)
#     print(Log_Path)
# except Exception as err:
#     print("Error occured while Log folder creation. Exception: ", err)
#     sys.exit()

try:
    audit_logger.info('Reading Config.ini file for SP_Password and Sp_Link.')
    parser = configparser.ConfigParser()
    parser.read(script_path + '\\config.ini')
    audit_logger.info("Loaded the config file")
    sp_username = parser.get('PATH', 'spUser')
    sp_password_token = parser.get('PATH', 'spPwd_token')
    leveredge_token = parser.get('PATH', 'leveredge_token')
    leveredge_details = common.login_details(token=leveredge_token)
    if not leveredge_details:
        raise Exception("leveredge credentials are not generated")
    leveredge_user = leveredge_details[0]
    leveredge_pass_1 = leveredge_details[1]
    leveredge_pass_2 = leveredge_details[2]
    sp_password_array = common.login_details(token=sp_password_token)
    if not sp_password_array:
        raise Exception("Sharepoint Password Credentials are not generated")
    sp_password = sp_password_array[0]
    mapping_fname = parser.get('PATH', 'mappingfileName')
    mapping_url = parser.get('PATH', 'mappingFilePath')
    to = parser.get('PATH', 'receipient')
    cc = parser.get('PATH', 'cc')
    subject = parser.get('PATH', 'subject')
    body = parser.get('PATH', 'body')
    lever_link = parser.get('PATH', 'leverEdge_Link')

except:
    error_logger.error('Error while reading bot_config.ini:', exc_info=True)
    audit_logger.info('Task Failed')
    sys.exit()

# Step 2: Killing the excel
try:
    audit_logger.info('Step 2 - Closing all applications like Excel, Outlook, Skype, Teams, Windows Explorer, SAP, Bex')
    common.kill_process(['excel.exe', 'EXCEL.EXE'])
except:
    error_logger.error("Step 2 - Closing all applications failed", exc_info=True)
    audit_logger.info("Task Failed")
    sys.exit()

try:
    # Step 6: Downloading Mapping File
    audit_logger.info('Step 6 : Downloading Mapping_Fin_TTS on Invoice_Task1')
    mapping_full_url = mapping_url + '/' + mapping_fname
    mapping_flpath = os.path.join(input_path, mapping_fname)
    if not common.share_point_download(sp_username, sp_password, mapping_full_url, mapping_flpath):
        error_logger.error(f'Failed downloading {mapping_full_url}', exc_info=True)
        audit_logger.info('Error while downloading mapping file')
        audit_logger.info('step 7 : sending email for mapping file not available')
        # Step 7: Sending email upon download failure
        common.send_email(sp_username, sp_password, to, cc, subject, body)
        audit_logger.info('Task Failed')
        sys.exit()
    # open, refresh and save mapping file
    common.open_refresh_save_xl([mapping_flpath])
    audit_logger.info(f"Mapping file {mapping_fname} downloaded")
except:
    error_logger.error("Task failed at step 6 - ",exc_info=True)
    audit_logger.info("Step - 6 Failed")
    sys.exit()

# Step 8: Process start email
try:
    audit_logger.info("step 8: send mail for process start")
    to, cc, subject, body = common.email_data(mapping_flpath, 'Email', 2)
    common.send_email(sp_username, sp_password, to, cc, subject, body)
    audit_logger.info(f"Process-Start mail sent to: {to} and cc: {cc}")
except:
    audit_logger.info("Step 8: Failed")
    error_logger.error('Sending start mail failed', exc_info=True)
    sys.exit()

# Step 9: 10 11 12
try:
    # Step 9: Fetch Input_download sheet
    download_df = pd.read_excel(mapping_flpath, sheet_name="Input_Download")
    missing_files = []
    ip_download_time = datetime.now()
    for iter in range(len(download_df)):
        audit_logger.info("Downloading files")
        fname = download_df.iloc[iter]['InputFileName']
        fpath = download_df.iloc[iter]['InputFilePath']
        local_path = os.path.join(input_path, fname)

        # Step 10 & 11: Started downloading and Saved the downloaded files in Input folder
        if not common.share_point_download(sp_username, sp_password, fpath + '/' + fname, local_path):
            mis_file = str(iter + 1) + "." + fname
            missing_files.append(mis_file)
            audit_logger.info(f"Template Download Failed - {fname}")
        else:
            audit_logger.info("Downloaded " + download_df.iloc[iter]["InputFileName"] + " from sharepoint" + download_df.iloc[iter]['"InputFilePath'])

    # Step 12: Sending the email if there is any failures at download
    if len(missing_files) > 0:
        audit_logger.info("Error while downloading, sending an notification mail")
        to, cc, subject, body = common.email_data(mapping_flpath, 'Email', 3)
        body = body.replace('<<List of files unavailable in sharepoint to be listed here>>', str(missing_files)[1:-1])
        common.send_email(sp_username, sp_password, to, cc, subject, body)
        sys.exit()
except:
    error_logger.error("Error while downloading", exc_info=True)
    audit_logger.info("Task Failed")
    sys.exit()

# Step 13: updating the Column C and D
try:
    Application = Dispatch("Excel.Application")
    Application.Visible = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Application.EnableEvents = False
    map_file_data = Application.Workbooks.Open(mapping_flpath)
    input_download_sheet = map_file_data.Worksheets('Input_Download')
    Application.Worksheets('Input_Download').Activate()
    row_count = 0
    while True:
        row_count += 1
        fname = input_download_sheet.Cells(row_count, 0).Value
        if fname:
            if fname in [missing_files]:
                input_download_sheet.Cells(row_count, 0).Value = 'Not Done'
            else:
                input_download_sheet.Cells(row_count, 2).Value = 'Done'

            input_download_sheet.Cells(row_count, 3).Value = datetime.strftime(ip_download_time, '%d_%m_%Y %H:%M:%S')
        else:
            break

    map_file_data.Save()
    map_file_data.Close(True)
except:
    error_logger.error("Error while updating the input_donload sheet", exc_info=True)
    audit_logger.info("Task Failed")
    sys.exit()

# Step 14: Refresh the B2P_Extractor page
common.open_refresh_save_xl([mapping_flpath])

