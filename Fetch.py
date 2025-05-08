import os
import sys
import time
import traceback

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.common.keys import Keys

import Secret
from cryptography.fernet import Fernet

import Configuration

import Authentication

from zipfile import ZipFile

import shutil

import logging

from datetime import datetime

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

    logging.basicConfig(filename='DEBUG.log', encoding='utf-8', level=logging.INFO)


def Page_Is_Loading(Session):
    while True:
        X = Session.execute_script("return document.readyState")
        if X == "complete":
            return True
        else:
            yield False


def LoginToPortal(Session):

    print("LoginToPortal")
    print("--------------------------------------------------")

    print("- authenticating/authorizing")

    Session.get('https://powerdms.com/ui/login.aspx?SiteID=ElPasoCSO&formsauth=true')
    time.sleep(Delay)
    while not Page_Is_Loading(Session):
        time.sleep(Delay)
        continue

    Wait = WebDriverWait(Session, 5)

    Element = Wait.until(expected_conditions.presence_of_element_located((By.ID, "UserIDInput")))
    Element.send_keys(Codec.decrypt(Configuration.UserIDInput).decode())

    Element = Wait.until(expected_conditions.presence_of_element_located((By.ID, "PasswordInput")))
    Element.send_keys(Codec.decrypt(Configuration.PasswordInput).decode())

    Element = Wait.until(expected_conditions.presence_of_element_located((By.ID, "userLogin_authenticator_powerDMSUserAuthenticator_powerDMSAuthenticator_LoginBtn")))
    Element.click()

    #print("- skip E-Mail verification")

    #Element = Wait.until(expected_conditions.presence_of_element_located((By.LINK_TEXT, "Continue to PowerDMS")))
    #Element.click()

    time.sleep(1)
    print()

def CreateDocumentsExport(Session):

    print("CreateDocumentsExport")
    print("--------------------------------------------------")

    print("- navigate to DocumentExport page")

    Session.get('https://powerdms.com/admin/Configuration/AdminMenu.aspx')
    time.sleep(Delay)
    while not Page_Is_Loading(Session):
        time.sleep(Delay)
        continue

    print("- filter folder")

    Session.get('https://powerdms.com/admin/Configuration/DocumentExport.aspx')
    time.sleep(Delay)
    while not Page_Is_Loading(Session):
        time.sleep(Delay)
        continue

    Wait = WebDriverWait(Session, 5)

    Element = Wait.until(expected_conditions.presence_of_element_located((By.ID, "ctl00_ctl00_pageBody_cphConfigurationContent_cbExportAsPdf")))
    Element.click()

    #Element = Wait.until(expected_conditions.presence_of_element_located((By.ID, "ctl00_ctl00_pageBody_cphConfigurationContent_rcbFolders")))
    Element = Wait.until(expected_conditions.presence_of_element_located((By.ID, "folders")))
    Element.click()
   
    #Element = Wait.until(expected_conditions.presence_of_element_located((By.NAME, "searchInput")))
    #Element.click()
    Element.send_keys(Codec.decrypt(Configuration.Folder).decode())   
    time.sleep(1)

    Element.send_keys(Keys.TAB)
    time.sleep(1)
    
    Element.send_keys(Keys.ENTER)
    time.sleep(1)
    
    #Element = Wait.until(expected_conditions.presence_of_element_located((By.CLASS_NAME, "powGlobalSearchInput_searchItem")))
    #Elements = Session.find_elements(By.CLASS_NAME,"powGlobalSearchInput_searchItem")
    #for Element in Elements:
    #    Title = Element.get_attribute("title")
    #    if Title == Codec.decrypt(Configuration.Folder).decode():
    #        Element.click()
    #        break

    print("- initialize export")
 
    Element = Wait.until(expected_conditions.presence_of_element_located((By.ID, "ctl00_ctl00_pageBody_cphConfigurationContent_btnExportDocuments")))
    Element.click()

    print("- accept warning")

    Alert = Wait.until(expected_conditions.alert_is_present())
    Alert.accept()
    time.sleep(Delay)
       
    print()


def CheckDocumentsExport(Session):

    print("CheckDocumentsExport")
    print("--------------------------------------------------")

    Wait = WebDriverWait(Session, 5)

    Ready = False

    while not Ready:

        print('- DocumentsExport.zip "IN PROGRESS..."')

        Session.get('https://powerdms.com/admin/Configuration/DocumentExport.aspx')
        while not Page_Is_Loading(Session):
            time.sleep(Delay)
            continue

        try:
            Element = Wait.until(expected_conditions.presence_of_element_located((By.ID, "ctl00_ctl00_pageBody_cphConfigurationContent_hlDocumentExportDownload")))
            if Element:
                Ready = True
        except Exception as E:
            pass

    print('- DocumentsExport.zip "COMPLETE"')

    print()


def DownloadDocumentsExport(Session):

    print("DownloadDocumentsExport")
    print("--------------------------------------------------")

    try:
        if os.path.isfile(Zip):
            os.remove(Zip)
    except Exception as E:
        pass

    Wait = WebDriverWait(Session, 5)

    Session.get('https://powerdms.com/Documents/10095/DocumentsExport.zip')
    while not Page_Is_Loading(Session):
        time.sleep(Delay)
        continue

    Ready = False

    while not Ready:

        print('- DocumentsExport.zip "DOWNLOADING..."')

        try:
            if os.path.isfile(Zip):
                Ready = True
                print('- DocumentsExport.zip "DOWNLOADED"')
        except Exception as E:
            pass

        time.sleep(Delay)

    print()


def ExtractDocumentsExport():

    print("ExtractDocumentsExport")
    print("--------------------------------------------------")

    while not os.path.exists(Stage):
        os.makedirs(Stage, 777)

    with ZipFile(Zip, 'r') as zipObj:
        print('- DocumentsExport.zip "EXTRACTING"')
        zipObj.extractall(path = Stage)
        print('- DocumentsExport.zip "EXTRACTED"')

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
                try:
                    Now = datetime.now()
                    shutil.move(Source, Target)
                    print(Target)
                except Exception as X:
                    TimeStamp = Now.strftime("%Y-%m-%d %H:%M:%S")
                    Message = TimeStamp + " " + str(X)
                    logging.info(Message)
                    print(Message)
                    pass
                    

        if len(os.listdir(Path)) == 0:
            os.rmdir(Path)

    print()
    print('- DocumentsExport.zip "MIGRATED"')

    time.sleep(Delay) 

    print()


def Delete_Folders_Starting_With(path, start_string):   
    try:
        for folder_name in os.listdir(path):
            folder_path = os.path.join(path, folder_name)
            if os.path.isdir(folder_path) and folder_name.startswith(start_string):
                shutil.rmtree(folder_path)
                print(f"Deleted folder: {folder_path}")
    except Exception as e:
        print(f"Exception: {e}")
        logging.info(f"Exception: {e}")


def Main():
    Clear()
    
    time.sleep(5)
    
    while True:

        try:
        
            print("Clean/Purge EDGE Residual Scope Folder(s)")
            print("--------------------------------------------------")
            
            Delete_Folders_Starting_With('C:\Program Files', 'scoped_dir')

            CWD = os.getcwd()

            Options = webdriver.EdgeOptions()
           
            Preferences = {"download.default_directory":CWD}
            Options.add_experimental_option("prefs", Preferences);
            Options.add_argument("--log-level=3")
            #Options.add_argument("headless")

            Session = webdriver.Edge(options=Options)

            Clear()

            LoginToPortal(Session)
            CreateDocumentsExport(Session)
            CheckDocumentsExport(Session)
            DownloadDocumentsExport(Session)

            Session.quit

            ExtractDocumentsExport()
            MigrateDocuments()         

            break

        except Exception as E:

            print(E)
            time.sleep(Delay * 17)

        finally:

            print()         
           
            print("Clean/Purge EDGE Residual Folder(s)")
            print("--------------------------------------------------")
            
            Delete_Folders_Starting_With('C:\Program Files', 'edge_BITS_')
            Delete_Folders_Starting_With('C:\Program Files', 'msedge_url_fetcher_')
            Delete_Folders_Starting_With('C:\Program Files', 'chrome_url_fetcher_')
            Delete_Folders_Starting_With('C:\Program Files', 'chrome_Unpacker_')

if __name__ == '__main__':

    Clear()

    Global()

    Main()