import os
import sys
import comtypes.client
from itertools import combinations
import pyautogui
import shutil


AttachToInstance = True

SpecifyPath = False

ProgramPath = 'C:\Program Files (x86)\Computers and Structures\SAP2000 20\SAP2000.exe'

APIPath ='D:\PROJECTS\minipro-S6\SAP_API\trail'


if not os.path.exists(APIPath):

        try:

            os.makedirs(APIPath)

        except OSError:

            pass


ModelPath = APIPath + os.sep + 'API_1-001.sdb'


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

SapModel = mySapObject.SapModel

filename=r"D:\PROJECTS\minipro-S6\SAP_API\structures\lattice02.sdb"
ret = SapModel.File.OpenFile(filename)

listOfDamagePoints=[182,430]#,193,213,438,230,227,446,255]
Failcond=1
InJoints=["121","124","129","133","137","145","143","148","152","156","160","164","168"]

#print([cls.__name__ for cls in comtypes.__subclasses__()])
def autoGui(case,count):
    pyautogui.click(535, 40,interval=1)#display
    pyautogui.click(535, 486,interval=1)#show plot
    pyautogui.click(639,483)
    pyautogui.keyDown('shift')
    pyautogui.press('down',13)
    pyautogui.keyUp('shift')  
    pyautogui.click(806,499,interval=1)#click add
    pyautogui.click(1268, 774,interval=1)#click display
    pyautogui.click(608, 330,interval=1)#click file
    pyautogui.click(654, 438,interval=1)#click print file
    pyautogui.write(r"D:\PROJECTS\minipro-S6\SAP_API\src\Automation\tempDatafiles\model{}{}.txt".format(case,count))
    pyautogui.press('enter')
    pyautogui.click(1203,725,interval=1)
    pyautogui.click(1298,810,interval=1)
    return


for idx,case in enumerate(['A','B','C']):
    comb = combinations(listOfDamagePoints, idx+1)
    count=0
    for combination in comb:
        for val in combination:
            ret = SapModel.FrameObj.SetMaterialOverwrite(str(val), "SteelEHalved")
        ret = SapModel.File.Save("D:\PROJECTS\minipro-S6\SAP_API\structures\ModelDataset_{}\model{}{}.sdb".format(case,case,count))
        ret = SapModel.Analyze.RunAnalysis()
        
        autoGui(case,count)
        count+=1
        ret=SapModel.SetModelIsLocked(False)
        for val in combination:
            ret=SapModel.FrameObj.SetMaterialOverwrite(str(val), "Steel")
    print("Number of combinations: ",count)


fromfolder=r'D:\PROJECTS\minipro-S6\SAP_API\src\Automation\tempDatafiles'
tofolder=r'D:\PROJECTS\minipro-S6\SAP_API\src\Automation\datafiles'
print("Moving files...")
for filename in os.listdir(fromfolder):
    print(filename)
    fromfilepath=os.path.join(fromfolder,filename)
    tofilepath=os.path.join(tofolder,filename)  
    if filename.endswith(".txt")==1:
        shutil.move(fromfilepath,tofilepath)







