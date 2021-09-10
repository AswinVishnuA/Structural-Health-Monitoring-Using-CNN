import os
import comtypes
import sys
# from Automation.datagen import AttachToInstance
from pywinauto import Desktop




# check whether there is a running Sap window open
def isSapOpen():
    windows = Desktop(backend="uia").windows()
    listOfWindows=[w.window_text() for w in windows]
    for win in listOfWindows:
        if(str(win).startswith("SAP")):
            return True
    return False

#creates a folder   
def createApiPath(path):

    if not os.path.exists(path):

        try:

            os.makedirs(path)
            pass

        except OSError:

            pass

ProgramPath = 'C:\Program Files (x86)\Computers and Structures\SAP2000 20\SAP2000.exe'

APIPath =r'D:\PROJECTS\minipro-S6\SAP_API\src\Automation\trialApi'

SpecifyPath = False

createApiPath(APIPath)

ModelPath = APIPath + os.sep + 'API_1-001.sdb'

AttachToInstance=isSapOpen()

print(AttachToInstance)

if AttachToInstance:

    try:
        mySapObject = comtypes.client.GetActiveObject("CSI.SAP2000.API.SapObject")
    except (OSError, comtypes.COMError):
        print("No running instance of the program found or failed to attach.")
        sys.exit(-1)
else:
    helper = comtypes.client.CreateObject('SAP2000v20.Helper')
    helper = helper.QueryInterface(comtypes.gen.SAP2000v20.cHelper)
    if SpecifyPath:

        try:
            mySapObject = helper.CreateObject(ProgramPath)
        except (OSError, comtypes.COMError):

            print("Cannot start a new instance of the program from " + ProgramPath)

            sys.exit(-1)
    else:

        try:
            mySapObject = helper.CreateObjectProgID("CSI.SAP2000.API.SapObject")

        except (OSError, comtypes.COMError):

            print("Cannot start a new instance of the program.")

            sys.exit(-1)
    mySapObject.ApplicationStart()

print("end is near")






























