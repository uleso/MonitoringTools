from datetime import *
import os
     



currentDaytime = datetime.today()
Today = currentDaytime.date()
desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
STRDAY = str(Today)
def addDatePNM(desktop):
    dirRenamelist = []
    pathPNM = os.path.join(os.path.join(desktop,"PNM"),STRDAY)
    # print(pathPNM)
    dirs = os.listdir(pathPNM)   
    for filename in dirs:
        dirRenamelist.append(os.path.join(pathPNM,filename))  
    return dirRenamelist    

def addDateMaint(desktop):
    
    dirRenamelist = []    
    pathMaint = os.path.join(os.path.join(desktop,"Maintenance Nodes"),STRDAY)
    # print(pathMaint)
    dirs = os.listdir(pathMaint)
    for filename in dirs:
        dirRenamelist.append(os.path.join(pathMaint,filename))
    
    return dirRenamelist   

def renamer(path):
    filesToRename = []
    for i in path:
        if(os.listdir(i) != []):
            filesToRename.append(os.listdir(i))
            for item in filesToRename:
                for photoname in item:
                    pathToRename = (os.path.join(i,photoname))    
                    newpath = os.path.join(i, (photoname[0:6] + "_"+ STRDAY +".PNG"))
                    os.rename(pathToRename,newpath)                   
                    print(f"{pathToRename} file changed ")
                    print(f"{newpath} new path")





renamedirPNM = addDatePNM(desktop)
renamer(renamedirPNM)

renamedirMaint = addDateMaint(desktop)
renamer(renamedirMaint)
# renamedirMaint = addDateMaint(desktop)
# renamer(renamedirMaint) 