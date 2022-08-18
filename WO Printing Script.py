import time
import PyPDF2
from PyPDF2 import PdfMerger

print("Hello...\n")

time.sleep(1)

print("To print all the lockout forms and inspection forms please follow the prompts.\nA PDF with everything will appear with the date and time as the title.\n")

shortcut = "*REMOVED FOR PRIVACY*"
shortcut2 = "*REMOVED FOR PRIVACY*"

numofWO = input("Please input how many work orders you would like to print: ")
StrapperUnload = input("Do you have Strapper or Unload Gripper WOs? (y/n) " )
if StrapperUnload == "y":
    whichone = input("Type 's' if you have strapper inspections, 'u' if you have unload gripper or 'b' if you have both: ")

    if whichone == 's':
        first = input("Which Strapper line is first, 1 or 2? ")
        if first == "2":
            shortcut = "*REMOVED FOR PRIVACY*"

    if whichone == 'u':
        first = input("Which Unload line is first, 1 or 2? ")
        if first == "2":
            shortcut2 = "*REMOVED FOR PRIVACY*"
    
    if whichone == 'b':
        first = input("Which Strapper line is first, 1 or 2? ")
        if first == "2":
            shortcut = "*REMOVED FOR PRIVACY*"

        first = input("Which Unload line is first, 1 or 2? ")
        if first == "2":
            shortcut2 = "*REMOVED FOR PRIVACY*"
       




emptying = 0
ToEmpty = []

Strapper = []
UnloadGripper = []

merger = PdfMerger()
for i in range(int(numofWO)):
    
# Below is a dictionary with all the WO titles and corresponding file location 

    fileDict = {
        *REMOVED FOR PRIVACY*
    }

    #So we can double check for emptying
    dehackdict = {

        #Dehack 2
        "Dehack 2 Inspection and Lubrication": "Elevator to strapper",
        "Unload Gripper Inspection": "Hack unloader/unload gripper",
        "Line 2 Receiving Table Inspection and Lubrication": "Receiving table",
        "Walking Beam Conveyor Inspection": "Walking Beam conveyor",
        "Strapper Inspections": "Strap station (including cross strapper)",

        #Dehack 1
        "Dehack 1 Inspection and Lubrication": "Elevator conveyor",
        "Packet Head Inspection": "Packet head, chains empty & elevator empty and up so millwrights don't fall down",
        "Wood Gripper Inspection 1": "Suction relieved (not holding wood/cardboard)",
        "Void Maker Spreader Assembly Inspection & Lubrication": "Void maker area"
    }

    #formatting + input
    currWO = input("%i. " %(i+1))

    if currWO in fileDict:

        #Emptying Check 
        if currWO in dehackdict:
            emptying = emptying + 1
            ToEmpty.append(dehackdict[currWO])

        #Adds current WO Form
        merger.append(fileDict[currWO])

        #Checks if even or odd (looks up page number)
        evenOdd = open(fileDict[currWO], 'rb')
        PyPDF2.PdfFileReader(evenOdd,False)
        readpdf = PyPDF2.PdfFileReader(evenOdd)
        totalpages = readpdf.numPages

        #Adds blank page if odd to prevent wrapping
        if (totalpages % 2 == 1):
            merger.append("BlankPage.pdf")

    #check if strapper or unload gripper as these have two options, line 1 or line 2
    if currWO == "Strapper Inspections":
        if shortcut == "*REMOVED FOR PRIVACY*":
            shortcut = "*REMOVED FOR PRIVACY*"
        else:
            shortcut = "*REMOVED FOR PRIVACY*"

    if currWO == "Unload Gripper Inspection":
        if shortcut2 == "*REMOVED FOR PRIVACY*":
            shortcut2 = "*REMOVED FOR PRIVACY*"
        else:
            shortcut2 = "*REMOVED FOR PRIVACY*"

    #Adds LOTO form and Blank Page
    merger.append("REALLOTOForm.pdf")
    merger.append("BlankPage.pdf")


print("Loading...\n")

merger.write("Work Order Inspection Forms.pdf")
merger.close()

if emptying != 0:
    print("Warning... Dehack emptying Required!!! \n\n")
    print("--------------------------------------------------------\n")
    for i in range(emptying):
        print(ToEmpty[i])
print("\n")

input("Click enter to exit...")
