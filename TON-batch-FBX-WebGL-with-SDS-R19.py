import c4d
from c4d import gui, documents, plugins, storage, utils
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

def exportFBX(filePath,doc):
    
    pluginID = 1026370
    plug = plugins.FindPlugin(pluginID, c4d.PLUGINTYPE_SCENESAVER)
    if plug is None:
        return

    op = {}
    # Send MSG_RETRIEVEPRIVATEDATA to export plugin
    if plug.Message(c4d.MSG_RETRIEVEPRIVATEDATA, op):

        if "imexporter" not in op:
            return
        
        # BaseList2D object stored in "imexporter" key hold the settings
        xExport = op["imexporter"]
        if xExport is None:
            return
        
        # Change export settings
        
        #unit_scale = c4d.UnitScaleData()
        #unit_scale.SetUnitScale(1.0, c4d.DOCUMENT_UNIT_CM)
        #xExport[c4d.FBXEXPORT_SCALE] = unit_scale
        
        
        
        xExport[c4d.FBXEXPORT_FBX_VERSION] = c4d.FBX_EXPORTVERSION_7_4_0
        xExport[c4d.FBXEXPORT_ASCII] = False
        xExport[c4d.FBXEXPORT_TEXTURES] = True
        xExport[c4d.FBXEXPORT_EMBED_TEXTURES] = False
        
        xExport[c4d.FBXEXPORT_SAVE_NORMALS] = True
        
        xExport[c4d.FBXEXPORT_TRIANGULATE] = False
         
        
        xExport[c4d.FBXEXPORT_LIGHTS] = False
        xExport[c4d.FBXEXPORT_CAMERAS] = False
        xExport[c4d.FBXEXPORT_SPLINES] = False
        xExport[c4d.FBXEXPORT_SDS] = False
         
        
        # Export the document
        if documents.SaveDocument(doc, filePath, c4d.SAVEDOCUMENTFLAGS_DONTADDTORECENTLIST, pluginID):
            print "Successfully exported: " + str(filePath)
        else:
            print "Export failed for: " +  str(filePath)

def main():
    
    c4d.CallCommand(12305, 12305) # Show Console...
    c4d.CallCommand(13957); # Clear console
    
    
    # ***************************************
    # SET UP SECTION
    rootDir = 'D:\\PROJEKTY\\TON\\FINAL_products\\Export35'
    destFolderPath = 'D:\\PROJEKTY\\TON\\FINAL_products\\configurator-FBX-Export35'

    suffix = "-source.c4d" # last part of source C4D file, which will be used to create FBX
    # ****************************************************
    
    nFolders = 0
    nFiles = 0
    
    
    
    if not os.path.exists(rootDir):
        print ("Root folder does not exist! Change root folder path in your script")
        ctypes.windll.user32.MessageBoxW(0, "Root folder does not exist! Change root folder path in your script", "Missing folder", 1)
        return
    
    if not os.path.exists(destFolderPath):
        #print(directory)
        os.makedirs(destFolderPath)
    

    
    
    

    doc = c4d.documents.GetActiveDocument()
    
    level1Dirs = getSubdirs(rootDir)
    for level1Dir in level1Dirs:
        nFolders += 1
        #print(level1Dir)
        #level2Dirs = getSubdirs(os.path.join(rootDir, level1Dir))
        #for level2Dir in level2Dirs:
            #print(level2Dir)
        dirPath = os.path.join(rootDir, level1Dir)
        xFiles = getFiles(dirPath)
        for xFileName in xFiles:
            xFilePath = os.path.join(rootDir,level1Dir,xFileName)
            filename, file_extension = os.path.splitext(xFileName)
            #print (xFilePath)
            if xFileName.endswith(suffix):
                nFiles += 1
                print "Converting file " + str(xFileName)
                print ""
            
                load = documents.LoadFile(xFilePath)
                
                if not load:
                    print ("File" + str(xFilePath) + " not loaded")
                
                doc = documents.GetActiveDocument()
                
                

                # ***************************************************************
                # Search for objects with SDS tag, apply SDS and make it editable
                
                arrPolyToSDS = []
                nObjects = 0
    
                allPolyObj = get_all_objects(doc.GetFirstObject(), lambda x: x.CheckType(c4d.Opolygon), [])
    
                for xObj in allPolyObj:
                    xTag=xObj.GetTag(c4d.Tsds)
            
                    if xTag != None:

                        xTag[c4d.SDSTAG_USE_SDS_STEPS] = True # set USE SDS objects Tags temporarily to true
                        
                        #xTag[c4d.SDSTAG_SUBRAY] = 1
                        #xTag[c4d.SDSTAG_SUBEDITOR] = 1
                        arrPolyToSDS.append(xObj)
                        nObjects += 1
                
                topObj = doc.GetFirstObject()
                objName = topObj.GetName()
            
                
                for x in range(0, nObjects):
                
                    # Create the SDS objects for each object with SDS tag and convert it to polygonal 
                    h = c4d.BaseObject(c4d.Osds) # Create new HyperNURBS
                    h[c4d.SDSOBJECT_SUBEDITOR_CM] = 1 # Set Editor subdivision to 1 (value does not matter as each object has its own SDS value set in SDS tag )
                    h[c4d.SDSOBJECT_SUBDIVIDE_UV] = c4d.SDSOBJECT_SUBDIVIDE_UV_EDGE #(value does not matter as each object has its own SDS value set in SDS tag )
                    h[c4d.SDSOBJECT_SUBRAY_CM] = 1 # Set Render subdivisions - (value does not matter as each object has its own SDS value set in SDS tag )
                    h.SetName('SDS')
                    doc.InsertObject(h)
            
                c4d.EventAdd()
                
                index = 0
                mySDS = get_all_objects(doc.GetFirstObject(), lambda x: x.CheckType(c4d.Osds), [])
                for sds in mySDS:
                    sds.InsertUnder(topObj)
                    xPolyObj = arrPolyToSDS[index]
                    sds.SetName(arrPolyToSDS[index].GetName())
                    xPolyObj.InsertUnder(sds)
                    index += 1
                    
                c4d.EventAdd()
                 
                for sds in mySDS:
                    doc.SetActiveObject(sds,c4d.SELECTION_NEW) 
                    
                    c4d.CallCommand(12236, 12236) # Make Editable
                    c4d.CallCommand(1019951, 1019951) # Delete Without Children
                    #print str(arrPolyToSDS[index].GetName())
                
                
                c4d.EventAdd()
                                
                # **********************************************************************************************
                
                
                
                
                
                # **********************************************************************************************
                #FBX Export
                
                newFileName = xFileName[:-len(suffix)] +  ".fbx" # remove suffix from the FBX file name
                newFileName = newFileName.replace("-", "") #delete - from the FBX name
                fpathNew = os.path.join (destFolderPath, newFileName)
                removeFile (fpathNew) #if file exists, delete it first
                exportFBX(fpathNew,doc)
                #print str("SDS objects: " + str(nObjects))
                #print str("                           ")
                
                c4d.documents.KillDocument(doc)
                # **********************************************************************************************
    

    print ("             ")
    print (" *** GOOD JOB *** ")
    print ("Folders checked:" + str(nFolders))
    print ("Files converted:" + str(nFiles))


if __name__=='__main__':
    main()