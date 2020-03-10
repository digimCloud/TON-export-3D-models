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
        
def removeFile(filePath):
    if os.path.exists(filePath):
        os.remove(filePath)

def get_all_objects(op, filter, output):
    while op:
        if filter(op):
            output.append(op)
        get_all_objects(op.GetDown(), filter, output)
        op = op.GetNext()
    return output

def main():
    
    c4d.CallCommand(12305, 12305) # Show Console...
    c4d.CallCommand(13957); # Clear console
    



    # *************************************************************
    # *************************************************************
    # SET UP SECTION

    rootDir = 'D:\\PROJEKTY\\TON\\FINAL_products\\Export35'
    #rootDir = 'E:\\PROJEKTY_E\\TON\\SDS\\models'
    suffix = "-source.c4d" # last part of source C4D file name
    arrMatNames = ["wm","wood solid", "wd","wood weneer", "f","fabrics","m","metal","x","misc","af","cane"]
    # ************************************************************
    # ************************************************************


    ver = int(c4d.GetC4DVersion())
    if ver > 19000:
        print ("ERROR: You run this script in C4D version " + str (ver) + ". Please use C4D version 18 as this script is saving C4D files!")
        #sys.exit(0)
        return

    if not os.path.exists(rootDir):
        print ("Root folder does not exist! Change root folder path in your script")
        ctypes.windll.user32.MessageBoxW(0, "Root folder does not exist! Change root folder path in your script", "Missing folder", 1)
        return

    nFiles = 0
    nFolders = 0
    lenSuffix = len(suffix)
    

    currDoc = c4d.documents.GetActiveDocument()
    
    #rootDir=c4d.storage.LoadDialog(c4d.FILESELECTTYPE_SCENES, "Please choose a folder", c4d.FILESELECT_DIRECTORY)
    
    
    #dirList=os.listdir(path)
    
    level1Dirs = getSubdirs(rootDir)
    for level1Dir in level1Dirs:
        nFolders += 1
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

                # remove all SDS object tags, rename objects and materials
                allPolyObj = get_all_objects(newDoc.GetFirstObject(), lambda x: x.CheckType(c4d.Opolygon), [])
    
                for xObj in allPolyObj:
                    # rename objects - remove material prefix
                    xObjName = xObj.GetName()
                    newObjName = xObjName[xObjName.index('_')+1:]
                    xObj.SetName(newObjName)
                    
                    # remove SDS tags
                    xTag=xObj.GetTag(c4d.Tsds)
                   
                    if xTag != None:
                        xTag.Remove()
                    
                # change material names
                active_mats = newDoc.GetMaterials()

                for mat in active_mats:
                    matName = mat[c4d.ID_BASELIST_NAME]
                    xMatName = ''.join([i for i in matName if not i.isdigit()]) #remove all digits from material name - w0m -> wm, f0 -> f etc.
                    if xMatName in arrMatNames:
                        index = arrMatNames.index(xMatName)
                        newMatName = arrMatNames[index+1]
                        mat[c4d.ID_BASELIST_NAME] = newMatName
                    else:
                        print ("Error: material " + xMatMame + " in file " + xFilerName + "does not have full name in script settings! Please edit script setup settings")
                
                #tex = mat[c4d.MATERIAL_COLOR_SHADER] # change channel ID here

                #if tex and tex.GetType() == c4d.Xbitmap: 
                    #path = tex[c4d.BITMAPSHADER_FILENAME]
                    #name = splitext(basename(path))[0] # to include extension use: name = basename(path) 
                    #doc.AddUndo(c4d.UNDOTYPE_CHANGE_SMALL, mat)
                    #mat[c4d.ID_BASELIST_NAME] = name

                

                #C4D Export - Low Poly
                destFileName = xFileName[:-lenSuffix] # remove source file suffix from the file name
                newFileName = destFileName +  "-lowpoly.c4d"
                fpathNew = os.path.join (destPath, newFileName)
                removeFile (fpathNew) #if file exists, delete it
                c4d.documents.SaveDocument(newDoc,fpathNew,c4d.SAVEDOCUMENTFLAGS_DONTADDTORECENTLIST,c4d.FORMAT_C4DEXPORT)
                
                #C4D Export - High Poly
                # Create TOP SDS object and put model as a child
                firstC4DObject = newDoc.GetFirstObject()

                h = c4d.BaseObject(c4d.Osds) # Create new HyperNURBS
                h[c4d.SDSOBJECT_SUBEDITOR_CM] = 2 # Set Editor subdivision to 1 
                h[c4d.SDSOBJECT_SUBDIVIDE_UV] = c4d.SDSOBJECT_SUBDIVIDE_UV_EDGE 
                h[c4d.SDSOBJECT_SUBRAY_CM] = 2 # Set Render subdivisions - 
                h.SetName('Subdivision Surface')
                newDoc.InsertObject(h)
                c4d.EventAdd()
                
                firstC4DObject.InsertUnder(h)

                destFileName = xFileName[:-lenSuffix] # remove source file suffix from the file name
                newFileName = destFileName +  "-highpoly.c4d"
                fpathNew = os.path.join (destPath, newFileName)
                removeFile (fpathNew) #if file exists, delete it
                c4d.documents.SaveDocument(newDoc,fpathNew,c4d.SAVEDOCUMENTFLAGS_DONTADDTORECENTLIST,c4d.FORMAT_C4DEXPORT)
                
                c4d.documents.KillDocument(newDoc)
           
            
    print ("Folders:  " + str(nFolders))
    print ("End of conversion")
    
    

        
        
    c4d.EventAdd()
    
    print "   "
    print "*** Good job! All files have been converted. ***"

if __name__=='__main__':
    main()