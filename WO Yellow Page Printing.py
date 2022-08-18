#https://stackoverflow.com/questions/29858752/error-message-chromedriver-executable-needs-to-be-available-in-the-path

from time import sleep
import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
import PyPDF2
from PyPDF2 import PdfMerger
from os import remove
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By



# Using Chrome to access web

print("Hello...")

numOfWO = int(input("How many WOs would you like to print today? "))
print('\n')

WOs = []

timeout = 5

for i in range(numOfWO):

    currWO = input("%i. " %(i+1))
    WOs.append(currWO)

chrome_options = webdriver.ChromeOptions()
prefs = {'download.default_directory' : 'Q:\\Maintenance\\Student files\\WO Printing\\DailyEAMDownloads\\'}
chrome_options.add_experimental_option('prefs', prefs)
driver = webdriver.Chrome(options=chrome_options)

for j in range(numOfWO):
    # Open the website
    driver.get('http://bbl-sql2/Reports/report/EAM/EAM-WorkOrders')

    sleep(5)


    driver.switch_to.frame(driver.find_element(By.XPATH, '//*[@id="main"]/div/paginated-report-viewer/div/iframe'))

    #driver.maximize_window() # For maximizing window
    #sleep(5)

    try:
        wo_box = driver.find_element(By.NAME, 'ReportViewerControl$ctl04$ctl03$txtValue')
        WebDriverWait(driver, timeout).until(wo_box)
    except TimeoutError:
        print("Timed out wait for page to load")
    wo_box.send_keys(WOs[j])

    try:
        report = driver.find_element(By.NAME, "ReportViewerControl$ctl04$ctl00")
        WebDriverWait(driver, timeout).until(report)
    except TimeoutException:
        print("Timed out wait for page to load")
    report.click()

    try:
        savebutton = driver.find_element(By.XPATH, '//*[@id="ReportViewerControl_ctl05_ctl04_ctl00_ButtonImgDown"]')
        WebDriverWait(driver, timeout).until(savebutton)
    except TimeoutException:
        print("Timed out wait for page to load")
    savebutton.click()

    try:
        pdfbutton = driver.find_element(By.XPATH, '//*[@id="ReportViewerControl_ctl05_ctl04_ctl00_Menu"]/div[4]/a')
        WebDriverWait(driver, timeout).until(pdfbutton)
    except TimeoutException:
        print("Timed out wait for page to load")
    pdfbutton.click()

    sleep(5)

    driver.close

driver.quit()


merger = PdfMerger()

for i in range(numOfWO):
    currentFile = "Q:/Maintenance/Student files/WO Printing/DailyEAMDownloads/EAM-WorkOrders (" + str(i + 1) + ").pdf"
    merger.append(currentFile)
    
    #Checks if even or odd (looks up page number)
    evenOdd = open(currentFile, 'rb')
    PyPDF2.PdfFileReader(evenOdd,False)
    readpdf = PyPDF2.PdfFileReader(evenOdd)
    totalpages = readpdf.numPages

    evenOdd.close()

    #Adds blank page if odd to prevent wrapping
    if (totalpages % 2 == 1):
        merger.append("BlankPage.pdf")
    


merger.write("Work Order Forms.pdf")
merger.close()

for i in range(numOfWO):
    currentFile = "Q:/Maintenance/Student files/WO Printing/DailyEAMDownloads/EAM-WorkOrders (" + str(i + 1) + ").pdf"
    remove(currentFile)
    

input("Click enter to exit...")