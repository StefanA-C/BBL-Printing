REQUIREMENTS

In order to use the WO printing python program you must:
1. Have python installed
2. Have the PyPDF2 library installed
3. Have selenium installed
4. Have the chromedriver executable installed
5. Have win32com installed
6. Have access to the Q-drive. I am assuming you do as you can read this message.
7. A couple hours to read through the code in case it breaks and you have to fix it (to do this you also need to install vscode or your fav IDE)

-> Instructions on how to do all this is below


Overall Printing Process:
1. Reads in the input document and saves all the info into arrays
2. Prompts for the run plan, saving the info as formatted HTML code for the email later
3. Open a chrome tab, navigate to the first WO and download the WO to a custom folder
4. Adds a page if odd page count and sends print request
5. Opens relevant inspection pdf and appends the required blank pages + LOTO forms
6. Saves and prints this new pdf
 --- You have just printed one WO ---
7. Repeat steps 3 to 6 for all the WOs
8. Prepares the email / title page with relevant information and opens the draft
9. Prepares emptying email (all dehack WOs get emptying) and opens draft if possible
10. Done


Once Installed:
1. Check if properly installed by clicking windows key-R
2. Type cmd.exe and press enter
3. Type "python --version" and enter
4. If anything other than something along the lines of "Python 3.10.4" appears you messed up, retry or google the resulting message
5. If it worked type "pip install PyPDF2" in the cmd (and press enter)
6. Type "pip install pywin32" into the cmd (and press enter)
8. Type "pip install selenium" into the cmd (and press enter)
9. Go to: https://chromedriver.chromium.org/downloads
10. Find what version of chrome you have (google how to find it) and download the webdriver that best matches it
11. Move the downloaded file (unzip first) to the scripts folder of the python version installed (for me it is: C:\Users\sarroyo-cottier\AppData\Local\Programs\Python\Python310\Scripts)
12. Open a pdf with adobe acrobat and ensure that it is set as the default to view all pdfs
13. Open any pdf with acrobat and select the print button, ensure the following options are selected
	- Choose paper source by pdf page size (selected)
	- Print on both sides of paper
14. Print the page to save the settings as the defaults (DO NOT CHANGE THESE AS THEY WILL MESS UP THE PRINTING PROCESS)
15. Log on to the outlook application so that the emails can be generated appropriately
16. Good to go
