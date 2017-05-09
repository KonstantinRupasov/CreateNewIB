# -*- coding: utf-8 -*-
"""
Create a new 1C infobase and publish it to IIS
    - 1: The name of the new IB (examle: my_new_ib)
"""
import pyodbc as db
import sys
import os
import shutil
import subprocess as sub
import re
from datetime import datetime


"""
Declare settings constants
Supposed to change rearely or never
Frequently changed settings are in command line parameters
"""
MSSQL_SERVER_NAME = 'NS3521530'
MSSQL_USER_NAME = 'sa'
MSSQL_USER_PASS = '14GiVv5S'
MSSQL_DB_FILES_PATH = 'C:\\Program Files\\Microsoft SQL Server\\MSSQL11.MSSQLSERVER\\MSSQL\\DATA\\'
MSSQL_BACKUP_PATH = 'C:\\Dropbox (1C-Poland)\\BACKUPS\\'
"""
TEMPLATE_NAME = 1C IB name
                IIS publication name
                SQL database name
Template SQL database:
    - is supposed to be DETACHED before the script is run
    - supposed to consist of 2 files in MSSQL_DB_FILES_PATH folder:
        - TEMPLATE_IB_NAME.mdf
        - TEMPLATE_IB_NAME_Log.ldf
"""
TEMPLATE_NAME = '1ctrade_template'   #1C IB name
TEMPLATE_USER = 'root'
TEMPLATE_PWD = 'root'
ONE_C_PATH = 'C:\\Program Files (x86)\\1cv8\\8.3.7.2027\\bin\\'
ONE_C_CREATE_IB_TEMPLATE_STR = '1cv8 CREATEINFOBASE Srvr=localhost;Ref={0};DBMS=MSSQLServer;DBSrvr={1};DBUID={2};DBPwd={3};DB={0}; /SLev1 /AddInList {0}'
IIS_APP_CMD = 'C:\\Windows\\System32\\inetsrv\\appcmd.exe'
IIS_SITE_NAME = 'Default Web Site'
WWW_ROOT_PATH = 'C:\\inetpub\\wwwroot\\'
LOG_2_FILE = False
LOG_PATH = 'C:\CreateNewIB\LOG'

def _exit(err_text, err_code):
    """
    Something went wrong
    Show error message and quit
    Error Codes:
        - ParametersError: something's wrong with the command line parameters
        - CannotConnectMSSQL: error connecting to MS SQL Server
        - FileNotFound: cannot find the file
        - FileAlreadyExists: file with the same name already exists
        - AppCmdError: Error running appcmd.exe
    """
    if LOG_2_FILE:
        print('**** ERROR ****', file=log)
        print(err_text, file=log)
        log.close()
    else:
        print('**** ERROR ****')
        print(err_text)
    sys.exit(err_code)
    
def _copy_file(src_file, dst_file):
    """
    Copy src file to dst file
    """
    if not os.path.isfile(src_file):
        #src file does not exist
        _exit('File (%s) not found' % src_file, 'FileNotFound')
    if os.path.isfile(dst_file):
        #dst file already exists
        _exit('File (%s) already exists' % dst_file, 'FileAlreadyExists')
    #Copy src file to dts file
    try:
        shutil.copy(src_file, dst_file)
    except Exception as exc:
        _exit('Error copying file:' + str(exc), 'ErrorCopyingFile')

def _log(messages):
    for message in messages:
        if LOG_2_FILE:
            print(message, file=log)
        else:
            print(message)
    if LOG_2_FILE:
        log.flush()
    
"""
Check the command line parameters
"""
if len(sys.argv) < 2:
    _exit('You need to specify the new IB name in the first command line parameter', 'ParametersError')
new_ib_name = sys.argv[1]
    
log = open(LOG_PATH + '\\' + new_ib_name + '.txt', 'a')
_log(['CREATING NEW INFOBASE', 'Started at' + str(datetime.now())])

"""
Copy template database files into new IB database files
"""
_copy_file(MSSQL_DB_FILES_PATH + TEMPLATE_NAME + '.mdf', MSSQL_DB_FILES_PATH + new_ib_name + '.mdf')
_copy_file(MSSQL_DB_FILES_PATH + TEMPLATE_NAME + '_Log.ldf', MSSQL_DB_FILES_PATH + new_ib_name + '_Log.ldf')
_log(['New IB database files are created:', 
      MSSQL_DB_FILES_PATH + new_ib_name + '.mdf',
      MSSQL_DB_FILES_PATH + new_ib_name])

"""
Connect to MS SQL Server (template databse)
"""
connection_str = 'DRIVER={SQL Server};SERVER=%s;DATABASE=%s;UID=%s;PWD=%s' % (MSSQL_SERVER_NAME, 'master', MSSQL_USER_NAME, MSSQL_USER_PASS)
try:
    connection = db.connect(connection_str)
except db.Error as err:
    _exit(err, 'CannotConnectMSSQL')
connection.autocommit = True
cursor = connection.cursor()

"""
Prepare the TSQL string to attach database
"""
sql_str = "CREATE DATABASE \"{}\" ON (FILENAME = '{}'), (FILENAME = '{}') FOR ATTACH".format(
        new_ib_name, 
        MSSQL_DB_FILES_PATH + new_ib_name + '.mdf', 
        MSSQL_DB_FILES_PATH + new_ib_name + '_Log.ldf')
_log(['-------------------------------------',
      'About to execute this TSQL statement to attach the new IB database:', sql_str])

"""
Attach new IB files to MS SQL
"""
try:
    cursor.execute(sql_str)
    cursor.commit()
except Exception as exc:
    _exit('Error copying file:' + str(exc), 'ErrorCopyingFile')

_log(['Database %s is attached sucessfully' % new_ib_name])

"""
Create a full database backup
in order for Backup jobs to be able to run
"""
_backup_path = MSSQL_BACKUP_PATH + new_ib_name
sql_str = "BACKUP DATABASE [{db_name}] TO  DISK = N'{backup_path}\\first_full_backup.bak' \
WITH  RETAINDAYS = 1, NOFORMAT, NOINIT, NAME = N'first_full_backup', SKIP, REWIND, NOUNLOAD,  \
STATS = 10".format(db_name=new_ib_name, backup_path=_backup_path)
_log(['-------------------------------------',
      'About to create the first full backup of the database:', sql_str])
try:
    os.mkdir(_backup_path)
    cursor.execute(sql_str)
    cursor.commit()
except Exception as exc:
    _exit('Error making the first full backup:' + str(exc), 'ErrorMakingFullBackup')

"""
Run ras in order to make possible the following actions (available only through rac):
    - Set secutiry profile for the infobase
    - Set "Allow license issuing by 1C:Enterprise Server"
"""
_log(['-------------------------------------',
      'Checking if RAS is running'])
res = sub.call('tasklist | findstr ras.exe', shell=True)    #Check if ras is running
if not res == 0:                                            #Ras is not running
    _log(['RAS is not found. Running RAS...'])
    res = sub.call('ras.exe cluster', shell=True)
    if not res == 0:                                            
        _exit('Cannot run ras.exe', 'AppCmdError')
_log(['Ras is running'])

"""
Create a new 1C IB
"""
os.chdir(ONE_C_PATH)
#Get cluster GUID
res = sub.check_output('rac.exe cluster list').decode('utf-8')
match = re.search(r'cluster\s*: ', res)
cluster_guid_pos = match.end()
cluster_guid = res[cluster_guid_pos:cluster_guid_pos+36]
_log(['cluster={}'.format(cluster_guid)])
#command = ONE_C_CREATE_IB_TEMPLATE_STR.format(new_ib_name, MSSQL_SERVER_NAME, MSSQL_USER_NAME, MSSQL_USER_PASS)
command = 'rac infobase \
--cluster={cluster} \
create --name={name} \
--dbms=MSSQLServer --db-server={db_server} \
--db-user={db_user} --db-pwd={db_pwd} \
--db-name={db_name} --locale=pl --date-offset=2000 --security-level=1 \
--license-distribution=allow'.format(name=new_ib_name,
                                     cluster=cluster_guid,
                                     db_server=MSSQL_SERVER_NAME, 
                                     db_user=MSSQL_USER_NAME, 
                                     db_pwd=MSSQL_USER_PASS,
                                     db_name=new_ib_name)
_log(['-------------------------------------', 
      'About to create a new 1C infobase with command:', command])
try:
    #os.system(command)
    res = sub.check_output(command)
    #Get infobase GUID
    infobase_guid = res[11:].decode('utf-8')            #res format is "infobase : XXXXXXXX"
except Exception as exc:
    _exit('Error creating 1C IB' + str(exc), 'ErrorCreatingIB')
_log([res, 'New 1C infobase {} is created'.format(new_ib_name), 'IB GUID={}'.format(infobase_guid)])

"""
Change IB security profile
"""
command = 'rac infobase --cluster={} \
update --infobase={} --infobase-user={} --infobase-pwd={} \
--security-profile-name=userbase --safe-mode-security-profile-name=userbase'.format(
        cluster_guid, 
        infobase_guid,
        TEMPLATE_USER,
        TEMPLATE_PWD)
_log(['-------------------------------------', 
      'About to run this commant in order to set up security profiles',
      command])
res = sub.check_output(command)
_log([res, 'Securty profiles are set up'])

"""
Publish the IB to IIS
"""
vrd_template = WWW_ROOT_PATH + TEMPLATE_NAME + '\\default.vrd'
command = 'webinst -publish -iis \
-wsdir {ib_name} -dir C:\inetpub\wwwroot\{ib_name} \
-connstr Srvr=localhost;Ref={ib_name} \
-descriptor {vrd_template}'. format(ib_name=new_ib_name, vrd_template=vrd_template)
_log(['-------------------------------------', 
      'About to run this commant in order to publish IB to IIS',
      command])
res = sub.check_output(command, shell=True)
_log([res, 'IB is published'])

_log(['All DONE'])
log.close()
