import os
import sys
import time
import traceback

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions

import Secret
from cryptography.fernet import Fernet

import Configuration

import requests
import Authentication

from zipfile import ZipFile

import shutil


def Clear():

    if os.name == 'nt':
        os.system('cls')
    else:
        os.system('clear')

    print("\x1b[?25h") # Show Cursor


def Global():

    global Slash
    if os.name == 'nt':
        Slash = '\\'
    else:
        Slash = '/'

    global Delay
    Delay = 5

    global Codec
    Codec = Fernet(Secret.Key)

    global Zip
    Zip = 'DocumentsExport.zip'

    global Stage
    Stage = 'Documents'


def Page_Is_Loading(Session):
    while True:
        X = Session.execute_script("return document.readyState")
        if X == "complete":
            return True
        else:
            yield False


def CreateDocumentsExport():

    Options = webdriver.EdgeOptions()
    Options.add_argument("headless")
    Options.add_argument("--log-level=3")

    Session = webdriver.Edge(options=Options)
    
    Clear()

    print("CreateDocumentsExport")
    print("--------------------------------------------------")

    Wait = WebDriverWait(Session, 15)

    print("- authentication in progress")

    Session.get('https://powerdms.com/ui/login.aspx?SiteID=ElPasoCSO&formsauth=true')
    time.sleep(Delay)

    while not Page_Is_Loading(Session):
        time.sleep(Delay)
        continue

    Element = Wait.until(expected_conditions.presence_of_element_located((By.ID, "UserIDInput")))
    Element.send_keys(Codec.decrypt(Configuration.UserIDInput).decode())

    Element = Wait.until(expected_conditions.presence_of_element_located((By.ID, "PasswordInput")))
    Element.send_keys(Codec.decrypt(Configuration.PasswordInput).decode())

    Element = Wait.until(expected_conditions.presence_of_element_located((By.ID, "userLogin_authenticator_powerDMSUserAuthenticator_powerDMSAuthenticator_LoginBtn")))
    Element.click()

    print("- skip E-Mail verification")

    Element = Wait.until(expected_conditions.presence_of_element_located((By.LINK_TEXT, "Continue to PowerDMS")))
    Element.click()

    print("- filter folder")

    Session.get('https://powerdms.com/admin/Configuration/DocumentExport.aspx')
    time.sleep(Delay)

    while not Page_Is_Loading(Session):
        time.sleep(Delay)
        continue

    Element = Wait.until(expected_conditions.presence_of_element_located((By.ID, "ctl00_ctl00_pageBody_cphConfigurationContent_cbExportAsPdf")))
    Element.click()

    Element = Wait.until(expected_conditions.presence_of_element_located((By.ID, "ctl00_ctl00_pageBody_cphConfigurationContent_rcbFolders")))
    Element.click()

    Element = Wait.until(expected_conditions.presence_of_element_located((By.NAME, "searchInput")))
    Element.click()
    Element.send_keys(Codec.decrypt(Configuration.Folder).decode())
    
    Element = Wait.until(expected_conditions.presence_of_element_located((By.CLASS_NAME, "powGlobalSearchInput_searchItem")))
    Elements = Session.find_elements(By.CLASS_NAME,"powGlobalSearchInput_searchItem")
    for Element in Elements:
        Title = Element.get_attribute("title")
        if Title == Codec.decrypt(Configuration.Folder).decode():
            Element.click()
            break

    print("- initialize export")
 
    Element = Wait.until(expected_conditions.presence_of_element_located((By.ID, "ctl00_ctl00_pageBody_cphConfigurationContent_btnExportDocuments")))
    Element.click()

    print("- accept warning")

    Alert = Wait.until(expected_conditions.alert_is_present())
    Alert.accept()
    time.sleep(Delay)  

    Session.quit

    print()


def CheckDocumentsExport(Session):

    print("CheckDocumentsExport")
    print("--------------------------------------------------")

    Headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36 Edg/105.0.1343.53'}

    Data_Authentication = {'userLogin$authenticator$powerDMSUserAuthenticator$powerDMSAuthenticator$UserIDInput': Codec.decrypt(Configuration.UserIDInput).decode(),'userLogin$authenticator$powerDMSUserAuthenticator$powerDMSAuthenticator$PasswordInput': Codec.decrypt(Configuration.PasswordInput).decode(),'__EVENTTARGET': 'userLogin$authenticator$powerDMSUserAuthenticator$powerDMSAuthenticator$LoginBtn','__VIEWSTATE': Authentication.VIEWSTATE,'__EVENTVALIDATION': Authentication.EVENTVALIDATION}

    print("- authentication in progress")

    Session.post('https://powerdms.com/ui/login.aspx', data=Data_Authentication, headers=Headers)

    Response = Session.get('https://powerdms.com/users/13720732')

    X = Response.text.split(',"email":"')
    Y = X[1].split('","givenName":')

    EMail = Y[0].lower()

    if EMail == Codec.decrypt(Configuration.EMail).decode().lower():

        print("- authentication process PASS")

        Status = ""

        while Status != "complete":

            Response = Session.get("https://powerdms.com/admin/Configuration/DocumentExport.aspx")

            time.sleep(5)

            try:
                X = Response.text.split('class="successTxt">')
                Y = X[1].split('</span>')
                Status = Y[0].lower()
            except:
                pass

            print("- DocumentsExport.zip file is INITIALIZING")

        print("- DocumentsExport.zip file is READY")

    else:

        print("- authentication process FAIL.")

    print()


def DownloadDocumentsExport(Session):

    print("DownloadDocumentsExport")
    print("--------------------------------------------------")

    print("\x1b[?25l") # Hide Cursor

    with open(Zip, 'wb') as File:

        Response = Session.get('https://powerdms.com/Documents/10095/DocumentsExport.zip', stream=True)
        Size = Response.headers.get('content-length')

        if Size is None:
            File.write(Response.content)
        else:
            Current = 0
            Size = int(Size)

            for Data in Response.iter_content(chunk_size=max(int(Size/1000), 1024*1024)):

                Current += len(Data)
                File.write(Data)
                Done = int(50*Current/Size)

                sys.stdout.write('\r{}{}'.format('█' * Done, '.' * (50-Done)))
                sys.stdout.flush()

    print("\x1b[?25h") # Show Cursor

    print()


def ExtractDocumentsExport():

    print("ExtractDocumentsExport")
    print("--------------------------------------------------")

    while not os.path.exists(Stage):
        os.makedirs(Stage, 777)

    with ZipFile(Zip, 'r') as zipObj:
        print('- DocumentsExport.zip extract is INITIALIZED')
        zipObj.extractall(path = Stage)
        print('- DocumentsExport.zip extract is COMPLETE')

    os.remove(Zip)

    print()


def MigrateDocuments():

    print("MigrateDocuments")
    print("--------------------------------------------------")

    for Path, Folders, Files in os.walk(Stage, topdown = False):

        for File in sorted(Files, reverse = True):

            Source = Path + Slash + File

            Target_Path = Codec.decrypt(Configuration.Path).decode()
            Target_File = File
            
            File_Prefix = File[0:3]

            if File_Prefix.isnumeric():
                Target_Path = Target_Path + Slash + File_Prefix
            
            Target = Target_Path + Slash + Target_File

            if not os.path.exists(Target_Path):
                os.makedirs(Target_Path, 777)

            if File.lower() == 'index.htm' and os.path.isfile(Source):
                os.remove(Source)
            else:
                shutil.move(Source, Target)
                print(Target)

        if len(os.listdir(Path)) == 0:
            os.rmdir(Path)

    print()
    print("- migrate complete")

    time.sleep(Delay) 

    print()


def Main():

    while True:

        try:

            CreateDocumentsExport()

            with requests.Session() as Session:

                CheckDocumentsExport(Session)
                DownloadDocumentsExport(Session)
                ExtractDocumentsExport()

            MigrateDocuments()

            break

        except Exception as E:

            print(E)
            time.sleep(Delay * 12)

        finally:

            print()


if __name__ == '__main__':

    Clear()

    Global()

    Main()