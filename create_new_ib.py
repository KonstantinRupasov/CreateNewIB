# -*- coding: utf-8 -*-
"""
Create a new 1C infobase and publish it to IIS
Command line parameters:
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
ONE_C_CREATE_IB_TEMPLATE_STR = '1cv8 CREATEINFOBASE Srvr=localhost;Ref={0};DBMS=MSSQLServer;DBSrvr={1};DBUID={2};DBPwd={3};DB={0}; /AddInList {0} /Out C:\\Rupasov\\CreateNewIB\\1c_log.txt'
IIS_APP_CMD = 'C:\\Windows\\System32\\inetsrv\\appcmd.exe'
IIS_SITE_NAME = 'Default Web Site'
WWW_ROOT_PATH = 'C:\\inetpub\\wwwroot\\'
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
    print('**** ERROR ****', file=log)
    print(err_text, file=log)
    log.close()
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
        shutil.copyfile(src_file, dst_file)
    except Exception as exc:
        _exit('Error copying file:' + str(exc), 'ErrorCopyingFile')

def _log(messages):
    for message in messages:
        print(message, file=log)
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
_log(['About to execute this TSQL statement to attach the new IB database:', sql_str])

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
Create a new 1C IB
"""
os.chdir(ONE_C_PATH)
command = ONE_C_CREATE_IB_TEMPLATE_STR.format(new_ib_name, MSSQL_SERVER_NAME, MSSQL_USER_NAME, MSSQL_USER_PASS)
_log(['About to create a new 1C infobase with command:', command])
log.flush()
try:
    os.system(command)
except Exception as exc:
    _exit('Error creating 1C IB' + str(exc), 'ErrorCreatingIB')
_log(['New 1C infobase {} is created'.format(new_ib_name)])
log.flush()

"""
Copy template IIS publication to new IB publication
"""
src = WWW_ROOT_PATH + TEMPLATE_NAME
dst = WWW_ROOT_PATH + new_ib_name
try:
    shutil.copytree(src, dst)
except FileNotFoundError:
    _exit('Cannot find the template publication directory: %s' % src, 'FileNotFound')
_log(['Template IIS publication is copied to', dst])
    
"""
Replace publication name and IB name n file default.vrd
"""
#Read the file into memory
dv_file_name = dst + '\default.vrd'
if not os.path.isfile(dv_file_name):
    _exit('File (%s) does not exist' % dv_file_name, 'FileNotFound')
dv_file = open(dv_file_name, 'r')
default_vrd_text = dv_file.read()
dv_file.close()
#Reopen the file to write
dv_file = open(dv_file_name, 'w')
#Replace publication name
source_str = 'base="/{}"'.format(TEMPLATE_NAME)
result_str = 'base="/{}"'.format(new_ib_name)
default_vrd_text = default_vrd_text.replace(source_str, result_str)
#Replace IB name
source_str = 'Ref=&quot;{}'.format(TEMPLATE_NAME)
result_str = 'Ref=&quot;{}'.format(new_ib_name)
default_vrd_text = default_vrd_text.replace(source_str, result_str)
dv_file.write(default_vrd_text)
dv_file.close()
_log(['defaul.vrd file is changed'])

"""
Add application to IIS
"""
command = '{} add app /site.name:"{}" /path:/{} /physicalPath:{}{}'.format(
        IIS_APP_CMD,
        IIS_SITE_NAME,
        new_ib_name,
        WWW_ROOT_PATH,
        new_ib_name)
_log(['About to create new IIS aplication. Command:', command])
try:
    os.system(command)
except Exception as exc:
    _exit('Cannot create ISS application', 'AppCmdError')
_log(['IIS app is created sucessfully'])

"""
Run ras in order to make possible the following actions (available only through rac):
    - Set secutiry profile for the infobase
    - Set "Allow license issuing by 1C:Enterprise Server"
"""
res = sub.call('tasklist | findstr ras.exe', shell=True)    #Check if ras is running
if not res == 0:                                            #Ras is not running
    _log(['Ras is not found. Running ras...'])
    res = sub.call('ras.exe cluster', shell=True)
    if not res == 0:                                            
        _exit('Cannot run ras.exe', 'AppCmdError')
_log(['Ras is running'])

"""
Change new 1C IB settings
"""
res = sub.check_output('rac.exe cluster list').decode('utf-8')
match = re.search(r'cluster\s*: ', res)
cluster_guid_pos = match.end()
cluster_guid = res[cluster_guid_pos:cluster_guid_pos+36]
_log(['cluster={}'.format(cluster_guid)])
res = sub.check_output('rac.exe infobase --cluster={} summary list'.format(cluster_guid)).decode('utf-8')
#print(res)
match = re.search(r'name\s*: {}'.format(new_ib_name), res)
infobase_giud_pos = match.start()-38
infobase_guid = res[infobase_giud_pos : infobase_giud_pos+36]
_log(['infobase={}'.format(infobase_guid)])
#print(infobase_guid)
command = 'rac infobase --cluster={} \
update --infobase={} --infobase-user={} --infobase-pwd={} \
--license-distribution=allow --security-profile-name=userbase'.format(
        cluster_guid, 
        infobase_guid,
        TEMPLATE_USER,
        TEMPLATE_PWD)
_log(['About to run this commant in order to turn on server license distribution and security profile',
      command])
res = sub.check_output(command)
_log(['All DONE'])
log.close()
