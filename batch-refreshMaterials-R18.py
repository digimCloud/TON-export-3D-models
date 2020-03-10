import c4d
from c4d import gui, documents, plugins, storage
import os
import sys


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
        

def main():
    
    # *********************************************************
    # SET UP SECTION
    suffix = "-source.c4d" # last part of source C4D file name
    # *********************************************************


    nFiles = 0
    
    c4d.CallCommand(12305, 12305) # Show Console...
    c4d.CallCommand(13957); # Clear console
    
    currDoc = c4d.documents.GetActiveDocument()
    
    rootDir=c4d.storage.LoadDialog(c4d.FILESELECTTYPE_SCENES, "Please choose a folder", c4d.FILESELECT_DIRECTORY)
    
    
    #dirList=os.listdir(path)
    
    level1Dirs = getSubdirs(rootDir)
    for level1Dir in level1Dirs:
        
        #print(level1Dir)
        destPath = os.path.join(rootDir, level1Dir)

        xFiles = getFiles(destPath)
        for xFileName in xFiles:
            xFilePath = os.path.join(rootDir,level1Dir,xFileName)
            filename, file_extension = os.path.splitext(xFilePath)
            #print (xFilePath)
            if xFileName.lower().endswith(suffix):
                print "Converting file " + str(xFileName)
                print ""
            
                load = documents.LoadFile(xFilePath)
                
                
                if not load:
                    print ("File" + str(xFilePath) + " not loaded")
                
                newDoc = documents.GetActiveDocument()
                c4d.CallCommand(12253, 12253) # Render All Materials
                c4d.CallCommand(12091, 12091) # Gouraud Shading
                c4d.CallCommand(12149, 12149) # Frame Default
                c4d.EventAdd()
                nFiles += 1
                #newDoc.SaveDocument()
                c4d.CallCommand(12098) # Save
                
                c4d.documents.KillDocument(newDoc)
           

    print "Files refreshed: " + str(nFiles)
    
    

        
        
    
    
    print "   "
    print "*** Good job! All files have been refreshed. ***"

if __name__=='__main__':
    main()