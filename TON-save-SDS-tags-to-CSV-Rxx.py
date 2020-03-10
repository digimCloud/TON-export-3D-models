import os
import sys
import ctypes  # An included library with Python install. 
import csv
import c4d
from c4d import gui, documents, plugins, storage, utils

# ABOUT THIS SCRIPT
# Reploaces texture files
# Goes thru all nested level 1 folders and replaces textures with different textures
# Go to SetUp section to set up root path


# EXPERIMENTAL - not used in tghis programme
def walklevel(some_dir, level):
    some_dir = some_dir.rstrip(os.path.sep)
    assert os.path.isdir(some_dir)
    num_sep = some_dir.count(os.path.sep)
    for root, dirs, files in os.walk(some_dir):
        yield root, dirs, files
        num_sep_this = root.count(os.path.sep)
        if num_sep + level <= num_sep_this:
            del dirs[:]


def get_all_objects(op, filter, output):
    while op:
        if filter(op):
            output.append(op)
        get_all_objects(op.GetDown(), filter, output)
        op = op.GetNext()
    return output

def getSubdirs(parentDir):
    subdirs = []
    dirList = next(os.walk(parentDir))[1]  # Walk only first level Folders
    for xDir in dirList:
        subdirs.append(xDir)
    return subdirs
        #print(fname)

def getFiles(parentDir):
    files = []
    fileList = next(os.walk(parentDir))[2]  # Walk only first level Files
    for fname in fileList:
        files.append(fname)
        #print (fname)
    return files
        #print(fname)    

def readLines(filePath):
    r = csv.reader(open(filePath, "rt", newline=''), dialect="excel")
    return [l for l in r]




def main():
    
    c4d.CallCommand(12305, 12305) # Show Console...
    c4d.CallCommand(13957); # Clear console
    
    
    
    
    # ***************************************
    # SET UP SECTION
    root3dDir = 'D:\\PROJEKTY\\TON\\FINAL_products\\Final_ALL'
    
    csvFilePath = 'D:\\PROJEKTY\\TON\\FINAL_products\\Final_ALL\\ton_SDS_180510.csv'
    suffix = "-source.c4d" # last part of source C4D file name
        
    # ***************************************
    
    
    
    
    
    
    nFolders = 0
    nFiles = 0
    arrCSVLines = []
    
    
    if not os.path.exists(root3dDir):
        print (root3dDir)
        print ("Root folder with 3D files does not exist! Change root folder path in your script")
        ctypes.windll.user32.MessageBoxW(0, "Root folder does not exist! Change root folder path in your script", "Missing folder", 1)
        sys.exit(0)
    
    
    
    
    
    level1Dirs = getSubdirs(root3dDir)
    
    for level1Dir in level1Dirs:
        nFolders += 1
        #print(level1Dir)
        #level2Dirs = getSubdirs(os.path.join(rootDir, level1Dir))
        #for level2Dir in level2Dirs:
            #print(level2Dir)
        dirPath = os.path.join(root3dDir, level1Dir)
        xFiles = getFiles(dirPath)
        for xFileName in xFiles:
            xFilePath = os.path.join(root3dDir,level1Dir,xFileName)
            filename, file_extension = os.path.splitext(xFilePath)
            #print (xFilePath)
            if xFileName.lower().endswith(suffix):         
                #open C4D file
                load = documents.LoadFile(xFilePath)
                    
                if not load:
                    print ("File" + str(xFilePath) + " not loaded")
                    
                doc = documents.GetActiveDocument()
                
                
                # ***************************************************************
                # Search for objects with SDS tag, read SDS value and save to to array
                    
                arrSDS = [] # save all SDS tags and its params to array
                arrSDS.append(level1Dir) # get folder name, where C4D is saved
                    
                allPolyObj = get_all_objects(doc.GetFirstObject(), lambda x: x.CheckType(c4d.Opolygon), [])
        
                for xObj in allPolyObj:
                    xTag=xObj.GetTag(c4d.Tsds)
                
                    if xTag != None:
                        arrSDS.append(xObj.GetName()) #append object name
                        arrSDS.append(xTag[c4d.SDSTAG_SUBRAY]) #append SDS value for rendering
                
                
                c4d.documents.KillDocument(doc) # close document as we do not need it anymore
                
                arrCSVLines.append(arrSDS) # save line with 
                
                #***************************************************************
                
                    
        # go thru array and save lines to CSV file           
        with open(csvFilePath,'wb') as resultFile:
            wr = csv.writer(resultFile, dialect='excel')
            for xLine in arrCSVLines:
                #print xLine
                wr.writerow(xLine)
    
    print ("              ")   



if __name__=='__main__':
    main()