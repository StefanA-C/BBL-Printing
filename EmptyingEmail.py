import win32com.client

fileobject = open("InputDoc.txt", 'r')
output = fileobject.read()
splitInfo = output.split("\n")

ArrayWithAllInfo = []
numOfWO = len(splitInfo)

runPlan = []


#for i in range(numOfWO):

 #   ArrayWithAllInfo.append(splitInfo[i].split("\t"))

#fileobject.close()


def EmpytingEmail(ArrayWithAllInfo):
    """
    Sends Empything Email
    """
    dehackdict = {

        "Dehack 2 Inspection and Lubrication": "Elevator to strapper",
        "Unload Gripper Inspection": "Hack unloader/unload gripper",
        "Line 2 Receiving Table Inspection and Lubrication": "Receiving table",
        "Walking Beam Conveyor Inspection": "Walking Beam conveyor",
        "Strapper Inspections": "Strap station (including cross strapper)",
        "Dehack 1 Inspection and Lubrication": "Elevator conveyor",
        "Packet Head Inspection": "Packet head, chains empty & elevator empty and up so millwrights don't fall down",
        "Wood Gripper Inspection 1": "Suction cups so suction is relieved (not holding wood/cardboard)",
        "Void Maker Spreader Assembly Inspection & Lubrication": "Void maker area"
    }
    
    Outlook = win32com.client.Dispatch("Outlook.Application")
    olNs = Outlook.GetNamespace("MAPI")
    Inbox = olNs.GetDefaultFolder(6)

    #dayte = str(ArrayWithAllInfo[0][9]).split('/')

    betterdayte = ArrayWithAllInfo[0][9]

    mail = Outlook.CreateItem(0)
    mail.To = '*REMOVED FOR PRIVACY*'
    mail.Subject = 'Emptying Request for ' + betterdayte

    rqst = ""



    for i in range(len(ArrayWithAllInfo)):
        if ArrayWithAllInfo[i][1][0] == "8" or (ArrayWithAllInfo[i][1][0] == "C" and ArrayWithAllInfo[i][1][1] == "8"):
            rqst1 = "<p>- " 
            rqst11 = ArrayWithAllInfo[i][3] 
            rqst2 = " (WO " + ArrayWithAllInfo[i][0]+ ")" 
            if ArrayWithAllInfo[i][3] in dehackdict:
                if ArrayWithAllInfo[i][1][0] == "C":
                    rqst3 = ", please empty the " + dehackdict[ArrayWithAllInfo[i][3]] +" on line 2.</p>"
                else:
                    rqst3 = ", please empty the " + dehackdict[ArrayWithAllInfo[i][3]] +" on line 1.</p>"
            else:
                if ArrayWithAllInfo[i][1][0] == "C":
                    rqst3 = ", please empty the !AREA! on line 2.</p>"
                else:
                    rqst3 = ", please empty the !AREA! on line 1.</p>"
            rqst1 = rqst1.lower()
            rqst3 = rqst3.lower()
            rqst11 = rqst11.capitalize()
            rqst = rqst + rqst1 + rqst11 + rqst2 + rqst3


    #mail.Body = 'Message body'
    mail.HTMLBody = """
    <p>Hello,</p>
    <p>Please let the operators know of the following dehack job(s) that are going to happen on """ + betterdayte + """:</p>
    """ + rqst + """
    <p>Thanks,</p>
    <p>Stefan Arroyo-Cottier</p>
    <p>Maintenance Co-op</p>
    <p>Brampton Brick</p>

    """ #this field is optional

    # To attach a file to the email (optional):
    if rqst != "":
        mail.Display(False)
        mail.Save()