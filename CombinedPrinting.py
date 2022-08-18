#https://stackoverflow.com/questions/29858752/error-message-chromedriver-executable-needs-to-be-available-in-the-path
from pickle import TRUE
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementNotInteractableException
from LocateFileDict import fileDict
from EmptyingEmail import EmpytingEmail
from EmailFunction import sendEmail
from time import sleep
import datetime
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import PyPDF2
from PyPDF2 import PdfMerger
import os




#This Section will Open the .txt file and store all the information on it in Arraywithalltheinfo
#WO, Equipment, Equipment Description, Description

fileobject = open("InputDoc.txt", 'r')
output = fileobject.read()
splitInfo = output.split("\n")

ArrayWithAllInfo = []
numOfWO = len(splitInfo)

for i in range(numOfWO):

    ArrayWithAllInfo.append(splitInfo[i].split("\t"))

fileobject.close()

runPlantxt = ""
points = input("How many runplan details would you like to enter? ")

if int(points) == 0:
    runPlantxt = """<p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;margin-left:.5in;text-indent:-.25in;'><span style="font-family:Symbol;">&middot;</span><span style='font-size:9px;font-family:"Times New Roman",serif;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</span>N/A</p>"""
else:    
    for i in range(int(points)):
        currPoint = input("%i. " %(i+1))
        runPlantxt = runPlantxt + """<p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;margin-left:.5in;text-indent:-.25in;'><span style="font-family:Symbol;">&middot;</span><span style='font-size:9px;font-family:"Times New Roman",serif;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</span>"""+ currPoint + """</p>"""

# Using Chrome to access web
#Opens up the web "browser and preps the download folder"
chrome_options = webdriver.ChromeOptions()
prefs = {'download.default_directory' : '*REMOVED FOR PRIVACY*'}
chrome_options.add_experimental_option('prefs', prefs)
driver = webdriver.Chrome(options=chrome_options)

#Access each WO from the website
for j in range(numOfWO):
    timep = datetime.datetime.now()
    try:
        os.remove("*REMOVED FOR PRIVACY*")
        break
    except FileNotFoundError:
        print("Error")

    timep = datetime.datetime.now()
    timep = int(timep.strftime('%Y%m%d'))

    merger = PdfMerger()
    # Open the website
    driver.get('*REMOVED FOR PRIVACY*')

    while True:
        try:
            driver.switch_to.frame(driver.find_element(By.XPATH, '//*[@id="main"]/div/paginated-report-viewer/div/iframe'))
            break
        except NoSuchElementException:
            print("Fail.")
    
    while True:
        try:
            wo_box = driver.find_element(By.NAME, 'ReportViewerControl$ctl04$ctl03$txtValue')
            wo_box.send_keys(ArrayWithAllInfo[j][0])
            break
        except NoSuchElementException:
            print("Fail.")
   
    while True:
        try:
            report = driver.find_element(By.NAME, "ReportViewerControl$ctl04$ctl00")
            report.click()
            break
        except NoSuchElementException:
            print("Fail.")

    while True:
        try:
            sleep(2)
            savebutton = driver.find_element(By.XPATH, '//*[@id="ReportViewerControl_ctl05_ctl04_ctl00_ButtonImgDown"]')
            savebutton.click()
            pdf = driver.find_element(By.XPATH, '//*[@id="ReportViewerControl_ctl05_ctl04_ctl00_Menu"]/div[4]/a')
            pdf.click()
            break
        except ElementNotInteractableException or NoSuchElementException:
            print("Fail.")

    time_to_wait = 100
    time_counter = 0
    while not os.path.exists("*REMOVED FOR PRIVACY*"):
        sleep(0.1)
        time_counter += 1
        if time_counter > time_to_wait:break

    print("done")

    while True:
        try:
            currentFile = "*REMOVED FOR PRIVACY*"
            merger.append(currentFile)
            break
        except PermissionError:
            print("oh no")
    
    #Checks if even or odd (looks up page number)
    evenOdd = open(currentFile, 'rb')
    PyPDF2.PdfFileReader(evenOdd,False)
    readpdf = PyPDF2.PdfFileReader(evenOdd)
    totalpages = readpdf.numPages

    evenOdd.close()
    
    #Adds blank page if odd to prevent wrapping
    if (totalpages % 2 == 1):
        merger.append("BlankA4.pdf")

    #creates the file that is now a single WO

    while True:
        try:
            merger.write("*REMOVED FOR PRIVACY*" + str(timep) +".pdf")
            merger.close()
            break
        except PermissionError:
            print("oh no")
    
    while True:
        try:
            os.startfile("*REMOVED FOR PRIVACY*" + str(timep) +".pdf", "print")
            break
        except PermissionError:
            print("oh no")
    
    
    driver.close
    while True:
        try:
            os.remove("*REMOVED FOR PRIVACY*")
            break
        except FileNotFoundError:
            print("Error")

    #WO is printed, time to print the LOTO and the Inspection form if there
    merger = PdfMerger()

    if ArrayWithAllInfo[j][3] in fileDict:
        merger.append(fileDict[ArrayWithAllInfo[j][3]])

        evenOdd = open(fileDict[ArrayWithAllInfo[j][3]], 'rb')
        PyPDF2.PdfFileReader(evenOdd,False)
        readpdf = PyPDF2.PdfFileReader(evenOdd)
        totalpages = readpdf.numPages

        evenOdd.close()

        #Adds blank page if odd to prevent wrapping
        if (totalpages % 2 == 1):
            merger.append("BlankPage.pdf")

    elif ArrayWithAllInfo[j][3] == "Strapper Inspections":
        if ArrayWithAllInfo[j][1] == "824":
            merger.append(fileDict["Strapper Inspections 1"])

            evenOdd = open(fileDict["Strapper Inspections 1"], 'rb')
            PyPDF2.PdfFileReader(evenOdd,False)
            readpdf = PyPDF2.PdfFileReader(evenOdd)
            totalpages = readpdf.numPages

            evenOdd.close()

            #Adds blank page if odd to prevent wrapping
            if (totalpages % 2 == 1):
                merger.append("BlankPage.pdf")

        else:
            merger.append(fileDict["Strapper Inspections 2"])

            evenOdd = open(fileDict["Strapper Inspections 2"], 'rb')
            PyPDF2.PdfFileReader(evenOdd,False)
            readpdf = PyPDF2.PdfFileReader(evenOdd)
            totalpages = readpdf.numPages

            evenOdd.close()

            #Adds blank page if odd to prevent wrapping
            if (totalpages % 2 == 1):
                merger.append("BlankPage.pdf")
    
    merger.append("REALLOTOForm.pdf")
    merger.append("BlankPage.pdf")

    #creates the file that is now a single WO

    while True:
        try:
            merger.write("*REMOVED FOR PRIVACY*" + str(timep) +".pdf")
            merger.close()
            break
        except PermissionError:
            print("oh no")
    
    while True:
        try:
            os.startfile("*REMOVED FOR PRIVACY*" + str(timep) +".pdf", "print")
            break
        except PermissionError:
            print("oh no")
    

driver.quit()

sendEmail(ArrayWithAllInfo, runPlantxt)
EmpytingEmail(ArrayWithAllInfo)
input("Click enter to exit...")