import win32com.client

fileobject = open("InputDoc.txt", 'r')
output = fileobject.read()
splitInfo = output.split("\n")

f = ArrayWithAllInfo = []
numOfWO = len(splitInfo)

for i in range(numOfWO):

    ArrayWithAllInfo.append(splitInfo[i].split("\t"))

fileobject.close()
"""
runPlantxt = ""
points = input("How many runplan details would you like to enter? ")"""

def sendEmail(ArrayWithAllInfo, runPlantxt):
    """
    Takes in whole array for one day and will email you (me) with the nicely formatted email that is ready to print.
    Needs import win32com.client?
    """
    namedict = {
        *REMOVED FOR PRIVACY*
    }

    if ArrayWithAllInfo[0][17] in namedict:
        name = namedict[ArrayWithAllInfo[0][17]]
    else:
        name = ArrayWithAllInfo[0][17]

    Outlook = win32com.client.Dispatch("Outlook.Application")
    olNs = Outlook.GetNamespace("MAPI")
    Inbox = olNs.GetDefaultFolder(6)

    #dayte = str(ArrayWithAllInfo[0][9]).split('/')

    betterdayte = ArrayWithAllInfo[0][9]


    mail = Outlook.CreateItem(0)
    mail.To = '*REMOVED FOR PRIVACY*'
    mail.CC = "*REMOVED FOR PRIVACY*"
    mail.Subject = 'Work Orders for ' + betterdayte + " - NIGHT SHIFT"

    table = ""
    #Table Creation!
    for i in range(len(ArrayWithAllInfo)):
        table = table + """<tr>
                <td style="width: 55pt;border: 1pt solid black;padding: 0in 5.4pt;height: 15pt;vertical-align: bottom;">
                    <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="font-size:13px;color:black;">"""+ArrayWithAllInfo[i][0]+"""</span></strong></p>
                </td>
                <td style="width: 50pt;border-top: 1pt solid black;border-right: 1pt solid black;border-bottom: 1pt solid black;border-image: initial;border-left: none;padding: 0in 5.4pt;height: 15pt;vertical-align: bottom;">
                    <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="font-size:13px;color:black;">"""+ArrayWithAllInfo[i][1]+"""</span></strong></p>
                </td>
                <td style="width: 101pt;border-top: 1pt solid black;border-right: 1pt solid black;border-bottom: 1pt solid black;border-image: initial;border-left: none;padding: 0in 5.4pt;height: 15pt;vertical-align: bottom;">
                    <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="font-size:13px;color:black;">"""+ArrayWithAllInfo[i][2]+"""</span></strong></p>
                </td>
                <td style="width: 192pt;border-top: 1pt solid black;border-right: 1pt solid black;border-bottom: 1pt solid black;border-image: initial;border-left: none;padding: 0in 5.4pt;height: 15pt;vertical-align: bottom;">
                    <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="font-size:13px;color:black;">"""+ArrayWithAllInfo[i][3]+"""</span></strong></p>
                </td>
            </tr>"""



    #mail.Body = 'Message body'
    mail.HTMLBody = """
    <p style='margin:0in;margin-bottom:.0001pt;font-size:20px;font-family:"Calibri",sans-serif;'><strong>READ ALL WORK ORDER COMMENTS AND INSTRUCTIONS THROUGHLY</strong></p>
    <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><strong>HEALTH &amp; SAFETY</strong></p>
    <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;margin-left:.5in;text-indent:-.25in;'><span style="font-family:Symbol;">&middot;</span><span style='font-size:9px;font-family:"Times New Roman",serif;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</span>Lockout / Tag out your equipment when performing maintenance.</p>
    <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;margin-left:.5in;text-indent:-.25in;'><span style="font-family:Symbol;">&middot;</span><span style='font-size:9px;font-family:"Times New Roman",serif;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</span>Wear your safety glasses and required PPE at ALL times.</p>
    <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;margin-left:.5in;text-indent:-.25in;'><span style="font-family:Symbol;">&middot;</span><span style='font-size:9px;font-family:"Times New Roman",serif;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</span>Wear a face mask if you cannot maintain physical distance from another worker</p>
    <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;margin-left:.5in;text-indent:-.25in;'><span style="font-family:Symbol;">&middot;</span><span style='font-size:9px;font-family:"Times New Roman",serif;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</span>Replace all guards as you complete a job</p>
    <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'>&nbsp;</p>
    <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><strong>HOUSEKEEPING</strong></p>
    <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;margin-left:.5in;text-indent:-.25in;'><span style="font-family:Symbol;">&middot;</span><span style='font-size:9px;font-family:"Times New Roman",serif;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</span>Clean up ALL spills.</p>
    <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;margin-left:.5in;text-indent:-.25in;'><span style="font-family:Symbol;">&middot;</span><span style='font-size:9px;font-family:"Times New Roman",serif;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</span>Pick up any debris or garbage and place in the appropriate container.</p>
    <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;margin-left:.5in;text-indent:-.25in;'><span style="font-family:Symbol;">&middot;</span><span style='font-size:9px;font-family:"Times New Roman",serif;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</span>Treat others as you would like to be treated.</p>
    <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;margin-left:.5in;text-indent:-.25in;'><span style="font-family:Symbol;">&middot;</span><span style='font-size:9px;font-family:"Times New Roman",serif;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</span>Leave the shop and oil room in the same or better condition than you found it.</p>
    <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'>&nbsp;</p>
    <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><strong>DUTIES/TASKS</strong></p>
    <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;margin-left:.5in;text-indent:-.25in;'><span style="font-family:Symbol;">&middot;</span><span style='font-size:9px;font-family:"Times New Roman",serif;'>&nbsp; &nbsp; &nbsp; &nbsp;&nbsp;</span>Please read the comments on assigned work orders</p>
    <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;margin-left:.5in;text-indent:-.25in;'><span style="font-family:Symbol;">&middot;</span><span style='font-size:9px;font-family:"Times New Roman",serif;'>&nbsp; &nbsp; &nbsp; &nbsp;&nbsp;</span>Please complete work orders assigned</p>
    <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;margin-left:.5in;text-indent:-.25in;'><span style="font-family:Symbol;">&middot;</span><span style='font-size:9px;font-family:"Times New Roman",serif;'>&nbsp; &nbsp; &nbsp; &nbsp;&nbsp;</span>Please provide support for all running equipment.</p>
    <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;text-indent:.25in;'><span style="font-family:Symbol;">&middot;</span><span style='font-size:9px;font-family:"Times New Roman",serif;'>&nbsp; &nbsp; &nbsp; &nbsp;&nbsp;</span><u>Complete all Work Orders accurately &ndash; record hours worked and comments about the work you</u> <u>completed</u></p>
    <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;text-indent:.25in;'><u><span style="text-decoration: none;">&nbsp;</span></u></p>
    <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><strong>RUN PLAN</strong></p>
    """ + runPlantxt + """
    <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'>&nbsp;</p>
    <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'>"""+ name + """ - Please Complete In Order</p>
    <table style="border: none;width:398.0pt;margin-left:-.4pt;border-collapse:collapse;">
        <tbody>
            <tr>
                <td style="width: 55pt;border: 1pt solid black;padding: 0in 5.4pt;height: 15pt;vertical-align: bottom;">
                    <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><strong><span style="font-size:13px;color:black;">Work Order</span></strong></p>
                </td>
                <td style="width: 50pt;border-top: 1pt solid black;border-right: 1pt solid black;border-bottom: 1pt solid black;border-image: initial;border-left: none;padding: 0in 5.4pt;height: 15pt;vertical-align: bottom;">
                    <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><strong><span style="font-size:13px;color:black;">Equipment</span></strong></p>
                </td>
                <td style="width: 101pt;border-top: 1pt solid black;border-right: 1pt solid black;border-bottom: 1pt solid black;border-image: initial;border-left: none;padding: 0in 5.4pt;height: 15pt;vertical-align: bottom;">
                    <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><strong><span style="font-size:13px;color:black;">Equipment Description</span></strong></p>
                </td>
                <td style="width: 192pt;border-top: 1pt solid black;border-right: 1pt solid black;border-bottom: 1pt solid black;border-image: initial;border-left: none;padding: 0in 5.4pt;height: 15pt;vertical-align: bottom;">
                    <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><strong><span style="font-size:13px;color:black;">Description</span></strong></p>
                </td>
            </tr>
            """ + table + """
            
        </tbody>
    </table>

    """ #this field is optional

    # To attach a file to the email (optional):

    mail.Display(False)
    mail.Save()
#sendEmail(ArrayWithAllInfo, runPlantxt)
