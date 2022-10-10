from datetime import *
import os

currentDaytime = datetime.today()
Today = currentDaytime.date()

desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')

def makeDir(filename,desktop,day):
    dir = os.path.join(desktop,filename)
    print(dir)
    if(os.path.exists(dir)):
        pass
    else:
        try:
            os.mkdir(dir)
        except FileNotFoundError:
            print("Dosya yolu ile ilgili sıkıntı var kıral")
        
    strday=str(day)
    print(strday)
    dir = os.path.join(dir,strday)
    
    print(dir)
    
    if(os.path.exists(dir)):
        pass
    if(os.path.exists(dir) == False):
        
        try:
            os.mkdir(dir)
        except FileNotFoundError:
            print("Dosya yolu ile ilgili sıkıntı var kıral")
    return dir
def nestedDir(filename,parent):
    if(os.path.exists(parent) == False):
        os.mkdir(parent)
    else:
        dir = os.path.join(parent,filename)
        os.mkdir(dir)

NCTpath = makeDir("NCT",desktop,Today)
nestedDir("Pre",NCTpath)
nestedDir("Post",NCTpath)
nestedDir("Rechecks",NCTpath)
Maintpath = makeDir("Maintenance Nodes",desktop,Today)
nestedDir("Splunk",Maintpath)
nestedDir("Adhoc",Maintpath)
PNMpath = makeDir("PNM",desktop,Today)
nestedDir("Adhoc",PNMpath)
nestedDir("Splunk",PNMpath)
nestedDir("DOCSIS",PNMpath)



