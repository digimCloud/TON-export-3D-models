import c4d
from c4d import gui, documents, plugins, storage
import os
import sys


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

def removeFile(filePath):
    if os.path.exists(filePath):
        os.remove(filePath)





def main():





    # *****************************************************
    # *****************************************************
    # SET UP SECTION

    rootDir = 'D:\\PROJEKTY\\TON\\FINAL_products\\Export35'
    #rootDir = 'E:\\PROJEKTY_E\\TON\\SDS\\models'
    suffix = "-lowpoly.c4d" # last part of source file name
    # *****************************************************
    # *****************************************************



    if not os.path.exists(rootDir):
        print ("Root folder does not exist! Change root folder path in your script")
        ctypes.windll.user32.MessageBoxW(0, "Root folder does not exist! Change root folder path in your script", "Missing folder", 1)
        return

    nFiles = 0
    nFolders = 0
    lenSuffix = len(suffix)

    c4d.CallCommand(12305, 12305) # Show Console...
    c4d.CallCommand(13957); # Clear console

    currDoc = c4d.documents.GetActiveDocument()

    #rootDir=c4d.storage.LoadDialog(c4d.FILESELECTTYPE_SCENES, "Please choose a folder", c4d.FILESELECT_DIRECTORY)


    #dirList=os.listdir(path)

    level1Dirs = getSubdirs(rootDir)
    for level1Dir in level1Dirs:
        nFolders += 1
        #print(level1Dir)
        destPath = os.path.join(rootDir, level1Dir)

        xFiles = getFiles(destPath)
        bFileFound = False
        
        for xFileName in xFiles:
            xFilePath = os.path.join(rootDir,level1Dir,xFileName)
            filename, file_extension = os.path.splitext(xFilePath)
            #print (xFilePath)
            if xFileName.lower().endswith(suffix):
                bFileFound = True
                newDoc = c4d.documents.BaseDocument()
               
                c4d.documents.InsertBaseDocument(newDoc)
                
                        
                # import without dialogs
                flagsImport = c4d.SCENEFILTER_OBJECTS |  c4d.SCENEFILTER_MATERIALS
                c4d.documents.MergeDocument(newDoc,xFilePath,flagsImport,None)
                
                c4d.EventAdd()

                print "Converting file " + str(xFileName) + "\r"

                DXFdocName = xFileName[:-lenSuffix] # remove source file suffix from the file name
                
                newDoc.SetDocumentName(DXFdocName)
                
                #test if DXF file already exists - if so, delete it
                newFileName = DXFdocName +  ".dxf"
                fpathNew = os.path.join (destPath, newFileName)
                removeFile (fpathNew) #if file exists, delete it

                
                #ExportDXF with materials
                c4d.CallCommand(60000, 15) # Export Filter



                nFiles += 1

                c4d.documents.KillDocument(newDoc)
                
        if not bFileFound:
            print (" ..-lowpoly.c4d file not found in folder: " + Level1Dir) 

    print ("                ")
    print ("*** GOOD JOB ***")
    print ("Folders checked:  " + str(nFolders))
    print ("Files converted:  " + str(nFiles))
    print ("End of conversion")

    c4d.EventAdd()


if __name__=='__main__':
    main()