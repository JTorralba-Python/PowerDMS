import os
import sys
import time
import traceback

import Secret
from cryptography.fernet import Fernet

import Configuration
import Payload_Authentication
import Payload_Export

import requests
from bs4 import BeautifulSoup
from zipfile import ZipFile

import shutil

def Global():

    global Slash

    if os.name == 'nt':
        Slash = '\\'
    else:
        Slash = '/'

def Clear():

    if os.name == 'nt':
        os.system('cls')
    else:
        os.system('clear')

def Download():

    Headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36 Edg/105.0.1343.53'
    }

    Data_Authentication = {
        'userLogin$authenticator$powerDMSUserAuthenticator$powerDMSAuthenticator$UserIDInput': Codec.decrypt(Configuration.UserIDInput).decode(),
        'userLogin$authenticator$powerDMSUserAuthenticator$powerDMSAuthenticator$PasswordInput': Codec.decrypt(Configuration.PasswordInput).decode(),
        '__EVENTTARGET': 'userLogin$authenticator$powerDMSUserAuthenticator$powerDMSAuthenticator$LoginBtn',
        '__VIEWSTATE': Payload_Authentication.VIEWSTATE,
        '__EVENTVALIDATION': Payload_Authentication.EVENTVALIDATION
    }

    Data_Export = {
        '__VIEWSTATE': Payload_Export.VIEWSTATE,
        '__EVENTVALIDATION': Payload_Export.EVENTVALIDATION,
        'ctl00$ctl00$pageBody$cphConfigurationContent$cbExportAsPdf': 'on',
        'ctl00$ctl00$pageBody$cphConfigurationContent$btnExportDocuments': 'Export+Documents'
    }

    if Codec.decrypt(Configuration.Folder).decode() != '':
        Data_Export['ctl00$ctl00$pageBody$cphConfigurationContent$rcbFolders$rcbFolders_selectedItemsJsonInput'] = '[{"breadcrumbs":[],"isSystem":false,"allowedContentTypes":["Document","Folder"],"description":null,"id":"402860","siteId":null,"name":"' + Codec.decrypt(Configuration.Folder).decode().replace(' ', '+') + '","objectType":"Folder","permission":"Edit","isEditable":true,"isEditableWithCascade":true,"_hasType":false}]'

    with requests.Session() as Session:

        print("Authentication process initialized.")
        print()

        Session.post('https://powerdms.com/ui/login.aspx', data=Data_Authentication, headers=Headers)

        My_Profile = Session.get('https://powerdms.com/users/13720732')
    
        X = My_Profile.text.split(',"email":"')
        Y = X[1].split('","givenName":')

        EMail = Y[0].lower()
    
        if EMail == Codec.decrypt(Configuration.EMail).decode():
            print('Authentication complete for "' + EMail + '".' )
        else:
            print('Authentication process has FAILED.')
            time.sleep(5)
            exit()
        print()

        print("Export process initialized.")
        print()

        Export_Create = Session.post('https://powerdms.com/admin/Configuration/DocumentExport.aspx', data=Data_Export, headers=Headers)
    
        Export_Status = ''
        while Export_Status != 'complete':
            Export_Complete = Session.get('https://powerdms.com/admin/Configuration/DocumentExport.aspx')
            time.sleep(5)
            try:
                X = Export_Complete.text.split('class="successTxt">')
                Y = X[1].split('</span>')
                Export_Status = Y[0].lower()
            except:
                pass
            print('DocumentsExport.zip is not available quite yet.')

        if Export_Status == 'complete':
            print()
            print('DocumentsExport.zip is available for download.')
            print()
            
            print('Download initialzed.')
            print()
            Progress_Bar(Session)
            print('Download complete.')
            print()

            with ZipFile(Zip_File, 'r') as zipObj:
                print('Extract process initialized.')
                zipObj.extractall(path = Staging_Folder)
                print('Extract complete.')
                print()

            try:
                os.remove(Zip_File)
                print('DocumentsExport.zip has been purged.')
                print()
            except:
                pass
        else:
            print('Something went wrong.')
            print()

def Progress_Bar(Session):
    print("\x1b[?25l") # Hide Cursor
    with open(Zip_File, 'wb') as Save_File:
        Response = Session.get('https://powerdms.com/Documents/10095/DocumentsExport.zip', stream=True)
        Total = Response.headers.get('content-length')

        if Total is None:
            Save_File.write(Response.content)
        else:
            Downloaded = 0
            Total = int(Total)
            for Data in Response.iter_content(chunk_size=max(int(Total/1000), 1024*1024)):
                Downloaded += len(Data)
                Save_File.write(Data)
                Done = int(50*Downloaded/Total)
                sys.stdout.write('\r{}{}'.format('█' * Done, '.' * (50-Done)))
                sys.stdout.flush()
    sys.stdout.write('\n\n')
    print("\x1b[?25h") # Show Cursor

def Process():

    print('Move process initialized.')
    print()

    for Path, Folders, Files in os.walk(Staging_Folder, topdown = False):

        for File in sorted(Files, reverse = True):

            Source = Path + Slash + File

            Target_Path = Codec.decrypt(Configuration.Path).decode()
            Target_File = File
            
            File_Prefix = File[0:3]
            if File_Prefix.isnumeric():
                Target_Path = Codec.decrypt(Configuration.Path).decode() + Slash + File_Prefix
            
            Target = Target_Path + Slash + Target_File

            if not os.path.exists(Target_Path):
                try:
                    os.makedirs(Target_Path, 777)
                except:
                    print(traceback.format_exc())

            try:
                if File.lower() == 'index.htm' and os.path.isfile(Source):
                    os.remove(Source)
                else:
                    shutil.move(Source, Target)
                    print(Target)
            except:
                print(traceback.format_exc())

        try:
            if len(os.listdir(Path)) == 0:
                os.rmdir(Path)
        except:
            pass

    print()
    print('Move complete.')

    time.sleep(5)

def Main():

    Download()
    Process()

if __name__ == '__main__':


    Clear()

    Global()

    global Zip_File
    Zip_File = 'DocumentsExport.zip'

    global Staging_Folder
    Staging_Folder = 'Documents'

    global Codec
    Codec = Fernet(Secret.Key)

    if not os.path.exists(Staging_Folder):
                try:
                    os.makedirs(Staging_Folder, 777)
                except:
                    print(traceback.format_exc())

    if Staging_Folder != None and os.path.exists(Staging_Folder):
        Main()
    else:
        print('STaging_Folder not defined or not found or invalid.')
