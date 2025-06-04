import os, sys, traceback
from os import path

import time
import datetime
from dateutil.relativedelta import relativedelta

#Module pour manipuler le navigateur
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# from selenium.webdriver.firefox.firefox_binary import FirefoxBinary

# monExecutable = FirefoxBinary('chemin/vers/l'exécutable/de/Firefox')
# browser = webdriver.Firefox(firefox_binary= monExecutable

# Module pour manipuler les fichiers excel
import pandas

# On récupère la date du mois dernier (fichier traité = m-1)
WorkingDate = datetime.date.today() - relativedelta(months=1)
Workingmonth = (f'{WorkingDate.month:02d}')

# On récupère l'année en cours pour poinbter vers le bon repertoire.
# Exemple: C:\Users\g67274\Documents\CAD\Stats\[ANNEE EN COURS]\Feedback
WorkingYear = WorkingDate.year
progression =""
scriptPath = os.getcwd()
ExcelsPath = path.join("C:\\Users\\g67274\\Documents\\CAD - DOSSIER STEF\\Stats",str(WorkingYear),"Feedback")
ExcelFilePath = path.join(ExcelsPath,Workingmonth +".xlsx")
SeleniumDriverEXE = path.join(scriptPath , "SeleniumDriverEXE")

def DataframeSourceCreate():
    DfSource = pandas.read_excel(ExcelFilePath,0)
    return DfSource

def Dic_NumDossier_Create(DfSource):
    # Initialisation du dictionnaire contenant les numéros de dossier Enecad, et le Nom de l'agent qui a traité le dossier.
    DictDossierEnecad = {}
    for keys in DfSource.iloc[:,3]:
        DictDossierEnecad.update({keys:None})
    return DictDossierEnecad

def Get_PrenomNom_Agent(msg_console):
    
    # Tant que tout les noms des agents ne sont pas récupérés on boucle sur ledictionnaire contenant les dossiers
    while None in DictDossierEnecad.values():
        Enecad_Base_url = "https://enecad.enedis.fr/"
        Xpath_Prenom = "//div[@class='col-md-3 text-xs-right']//*/span[1]"
        Xpath_Nom = "//div[@class='col-md-3 text-xs-right']//*/span[2]"
        Xpath_NumDossier = "//div[@class='col-md-6 text-xs-center']//*/span[2]"

        msg_console = f"{msg_console} \nConnexion à Enecad"
        os.system('cls')
        print(msg_console)
        driver = webdriver.Firefox()
        driver.get(Enecad_Base_url)
        time.sleep(1)
        wait = WebDriverWait(driver, 60)
        wait.until(EC.url_to_be("https://enecad.enedis.fr/#/dashboard"))

        try:
            # On parcourt l'ensemble des numéros de dossiers du dictionnaire DictDossierEnecad
            for num_dossier in DictDossierEnecad.keys():
                progression = f"\nTraitement de l'enregistrement {list(DictDossierEnecad).index(num_dossier) + 1} sur {len(DictDossierEnecad)}"
                os.system('cls')
                print(msg_console, progression)

                # On recherche le premier dossier qui n'est pas rattaché à un agent.
                if DictDossierEnecad.get(num_dossier) == None:
                    Dossier_url = "https://enecad.enedis.fr/#/" + num_dossier
                    driver.get(Dossier_url)
                    wait = WebDriverWait(driver, 60)
                    wait.until(EC.text_to_be_present_in_element((By.XPATH,Xpath_NumDossier),num_dossier))

                    Prenom = driver.find_element(By.XPATH,Xpath_Prenom)
                    Prenom = Prenom.get_attribute("innerHTML")
                    Nom = driver.find_element(By.XPATH, Xpath_Nom)
                    Nom = Nom.get_attribute("innerHTML")
                    DictDossierEnecad.update({num_dossier:Nom + Prenom})
                    Prenom = ""
                    Nom = ""
            driver.close()
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            tb = traceback.extract_tb(exc_tb)[-1]
            print(exc_type, tb[2], tb[1], exc_obj)
            print("Erreur suivante rencontrée:"+ str(e))
            pass
    msg_console = f"{msg_console} {progression}"
    return msg_console

def DataframeUpdate():
    DfSource.insert(loc=2,column="Agent", value=DictDossierEnecad.values())

def ExcelUpdate():
    DfSource.to_excel(ExcelFilePath,index=False)

DfSource = DataframeSourceCreate()
DictDossierEnecad = Dic_NumDossier_Create(DfSource)

heure_debut = datetime.datetime.today()
msg_console = f"Heure de début du traitement: {heure_debut} \nLe fichier à traiter est le suivant: {ExcelFilePath} \nIl y a {len(DictDossierEnecad)} dossier(s) a traiter"
print(msg_console)

msg_console = Get_PrenomNom_Agent(msg_console)

DataframeUpdate()
msg_console = f"{msg_console} \nAjout de la colonne Agent et mise à jour du fichier"
os.system('cls')
print(msg_console)
ExcelUpdate()

msg_console = f"{msg_console} \nfin du traitement: {datetime.datetime.today()}"
os.system('cls')
print(msg_console) 
input("Traitement terminé, appuyez sur entrée pour quitter")