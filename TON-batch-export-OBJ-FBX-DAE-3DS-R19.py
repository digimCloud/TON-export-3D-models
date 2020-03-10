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


def exportABC(filePath,doc):
    # Get Alembic export plugin, 1028082 is its ID
    plug = plugins.FindPlugin(1028082, c4d.PLUGINTYPE_SCENESAVER)
    if plug is None:
        return
    
    # Get a path to save the exported file
    #filePath ="/Users/digitalmedia/Documents/TEMP/TEST-EXPORT/boxX.fbx"

    op = {}
    # Send MSG_RETRIEVEPRIVATEDATA to Alembic export plugin
    if plug.Message(c4d.MSG_RETRIEVEPRIVATEDATA, op):

        if "imexporter" not in op:
            return
        
        # BaseList2D object stored in "imexporter" key hold the settings
        abcExport = op["imexporter"]
        if abcExport is None:
            return
        
        #print 'SELECTION BEFORE: ',  abcExport[c4d.ABCEXPORT_SELECTION_ONLY]
        
        # Change Alembic export settings
        abcExport[c4d.ABCEXPORT_SELECTION_ONLY] = False
        abcExport[c4d.ABCEXPORT_PARTICLES] = False
        abcExport[c4d.ABCEXPORT_PARTICLE_GEOMETRY] = False
        
        
        #print 'SELECTION AFTER: ',  abcExport[c4d.ABCEXPORT_SELECTION_ONLY]


        # Finally export the document
        if documents.SaveDocument(doc, filePath, c4d.SAVEDOCUMENTFLAGS_DONTADDTORECENTLIST, 1028082):
            #print "Successfully exported: " + str(filePath)
            return
        else:
            print "Export failed!"


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
            #print "Successfully exported: " + str(filePath)
            return
        else:
            print "Export failed for: " +  str(filePath)



def exportOBJ(filePath,doc):
    
    pluginID = 1030178
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
        
        #unit_scale = c4d.UnitScaleData()
        #unit_scale.SetUnitScale(1.0, c4d.DOCUMENT_UNIT_CM)
        #xExport[c4d.OBJEXPORTOPTIONS_SCALE] = unit_scale
        
        xExport[c4d.OBJEXPORTOPTIONS_NORMALS] = True
        xExport[c4d.OBJEXPORTOPTIONS_MATERIAL] = True
        
        xExport[c4d.OBJEXPORTOPTIONS_MATERIAL] = c4d.OBJEXPORTOPTIONS_MATERIAL_MATERIAL
        xExport[c4d.OBJEXPORTOPTIONS_TEXTURECOORDINATES] = True
        xExport[c4d.OBJEXPORTOPTIONS_OBJECTS_AS_GROUPS] = True
        
        xExport[c4d.OBJEXPORTOPTIONS_POINTTRANSFORM_FLIPX] = False
        xExport[c4d.OBJEXPORTOPTIONS_POINTTRANSFORM_FLIPY] = False
        xExport[c4d.OBJEXPORTOPTIONS_POINTTRANSFORM_FLIPZ] = False
        
        xExport[c4d.OBJEXPORTOPTIONS_POINTTRANSFORM_SWAPYZ] = False
        xExport[c4d.OBJEXPORTOPTIONS_POINTTRANSFORM_SWAPXZ] = False
        xExport[c4d.OBJEXPORTOPTIONS_POINTTRANSFORM_SWAPXY] = False
        
        xExport[c4d.OBJEXPORTOPTIONS_INVERT_TRANSPARENCY] = False
        
        xExport[c4d.OBJEXPORTOPTIONS_FLIPFACES] = False
        xExport[c4d.OBJEXPORTOPTIONS_FLIPUVW] = False

        # Export the document
        if documents.SaveDocument(doc, filePath, c4d.SAVEDOCUMENTFLAGS_DONTADDTORECENTLIST, pluginID):
            #print "Successfully exported: " + str(filePath)
            return
        else:
            print "Export failed for: " +  str(filePath)



def exportDAE14(filePath,doc):
    
    pluginID = 1022316
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
        
        #unit_scale = c4d.UnitScaleData()
        #unit_scale.SetUnitScale(1.0, c4d.DOCUMENT_UNIT_CM)
        #xExport[c4d.COLLADA_EXPORT_SCALE] = unit_scale
        
        # Change export settings
        xExport[c4d.COLLADA_EXPORT_TRIANGLES] = False
        xExport[c4d.COLLADA_EXPORT_ANIMATION] = False
        xExport[c4d.COLLADA_EXPORT_GROUP] = True
        xExport[c4d.COLLADA_EXPORT_2D_GEOMETRY] = False

        
        # Export the document
        if documents.SaveDocument(doc, filePath, c4d.SAVEDOCUMENTFLAGS_DONTADDTORECENTLIST, pluginID):
            #print "Successfully exported: " + str(filePath)
            return
        else:
            print "Export failed for: " +  str(filePath)


def exportDAE15(filePath,doc):
    
    pluginID = 1025755
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
        xExport[c4d.COLLADA_EXPORT_TRIANGLES] = False
        xExport[c4d.COLLADA_EXPORT_ANIMATION] = False
        xExport[c4d.COLLADA_EXPORT_GROUP] = True
        xExport[c4d.COLLADA_EXPORT_2D_GEOMETRY] = False
        xExport[c4d.COLLADA_EXPORT_SCALE] = 1
        
        # Export the document
        if documents.SaveDocument(doc, filePath, c4d.SAVEDOCUMENTFLAGS_DONTADDTORECENTLIST, pluginID):
            #print "Successfully exported: " + str(filePath)
            return
        else:
            print "Export failed for: " +  str(filePath)


def export3DS(filePath,doc):
    
    pluginID = 1001038
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
        #xExport[c4d.F3DEXPORTOPTIONS_SCALE] = unit_scale
        
        xExport[c4d.F3DSEXPORTFILTER_GROUP] = False
        xExport[c4d.F3DSEXPORTFILTER_SPLITTEXCOORDS] = True
        

        
        # Export the document
        if documents.SaveDocument(doc, filePath, c4d.SAVEDOCUMENTFLAGS_DONTADDTORECENTLIST, pluginID):
            #print "Successfully exported: " + str(filePath)
            return
        else:
            print "Export failed for: " +  str(filePath)


def exportVRML2(filePath,doc):
    
    pluginID = 1001034
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
        
        #unit_scale = c4d.UnitScaleData()
        #unit_scale.SetUnitScale(1.0, c4d.DOCUMENT_UNIT_CM)
        #xExport[c4d.VRML2EXPORTFILTER_SCALE] = unit_scale
        
        xExport[c4d.OBJEXPORTOPTIONS_NORMALS] = True
        
        xExport[c4d.VRML2EXPORTFILTER_TEXTURES] = c4d.VRML2EXPORTFILTER_TEXTURES_REFERENCED   # VRML2EXPORTFILTER_TEXTURES_NONE, VRML2EXPORTFILTER_TEXTURES_INFILE
        xExport[c4d.VRML2EXPORTFILTER_BACKFACECULLING] = True
        xExport[c4d.VRML2EXPORTFILTER_FORMAT] = True
        
        #xExport[c4d.VRML2EXPORTFILTER_TEXTURESIZE] = 
        
        #xExport[c4d.VRML2EXPORTFILTER_SAVEANIMATION] = False
        #Export[c4d.VRML2EXPORTFILTER_KEYSPERSECOND] = 25

        # Export the document
        if documents.SaveDocument(doc, filePath, c4d.SAVEDOCUMENTFLAGS_DONTADDTORECENTLIST, pluginID):
             #print "Successfully exported: " + str(filePath)
             return
        else:
            print "Export failed for: " +  str(filePath)





def main():
    




    # *****************************************************
    # *****************************************************
    # SET UP SECTION

    rootDir = 'D:\\PROJEKTY\\TON\\FINAL_products\\Export34'
    #rootDir = 'E:\\PROJEKTY_E\\TON\\SDS\\models'
    suffix = "-source.c4d" # last part of source C4D file name
    arrMatNames = ["wm","wood solid","wd","wood weneer","f","fabrics","m","metal","x","misc","af","cane"]
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
        for xFileName in xFiles:
            xFilePath = os.path.join(rootDir,level1Dir,xFileName)
            filename, file_extension = os.path.splitext(xFilePath)
            #print (xFilePath)
            if xFileName.lower().endswith(suffix):
            
                print "Converting file " + str(xFileName) + "\r"
            
                load = documents.LoadFile(xFilePath)
                
                if not load:
                    print ("File" + str(xFilePath) + " not loaded")
                
                newDoc = documents.GetActiveDocument()
                
                allPolyObj = get_all_objects(newDoc.GetFirstObject(), lambda x: x.CheckType(c4d.Opolygon), [])

                # remove object name material prefixes
                for xObj in allPolyObj:
                    # rename objects - remove material prefix
                    xObjName = xObj.GetName()
                    newObjName = xObjName[xObjName.index('_')+1:]
                    print (newObjName)
                    xObj.SetName(newObjName)
                    
                    # remove SDS tags
                    xTag=xObj.GetTag(c4d.Tsds)
                    if xTag != None:
                        xTag.Remove()
                    
                # change material names
                active_mats = newDoc.GetMaterials()

                for mat in active_mats:
                    matName = mat[c4d.ID_BASELIST_NAME]
                    print (matName)
                    xMatName = ''.join([i for i in matName if not i.isdigit()]) #remove all digits from material name - w0m -> wm, f0 -> f etc.
                    if xMatName in arrMatNames:
                        index = arrMatNames.index(xMatName)
                        newMatName = arrMatNames[index+1]
                        mat[c4d.ID_BASELIST_NAME] = newMatName
                    else:
                       print ("Error: material " + xMatMame + " in file " + xFilerName + "does not have full name in script settings! Please edit script setup settings")
                
                destFileName = xFileName[:-lenSuffix] # remove source file suffix from the file name

                #FBX Export
                newFileName = destFileName +  ".fbx" 
                fpathNew = os.path.join (destPath, newFileName)
                removeFile (fpathNew) #if file exists, delete it
                exportFBX(fpathNew,newDoc)
                
                #OBJ Export
                newFileName = destFileName +  ".obj" 
                newFileNameMTL = destFileName +  ".mtl"
                fpathNew = os.path.join (destPath, newFileName)
                fpathNewMTL = os.path.join (destPath, newFileNameMTL)
                removeFile (fpathNew) #if file exists, delete it
                removeFile (fpathNewMTL) #if MTL file exists, delete it
                exportOBJ(fpathNew,newDoc)
                
                
                #Collada 1.4 Export
                newFileName = destFileName +  ".dae" 
                fpathNew = os.path.join (destPath, newFileName)
                removeFile (fpathNew) #if file exists, delete it
                exportDAE14(fpathNew,newDoc)
                
                #Collada 1.5 Export
                #newFileName = destFileName + "-collada1_5" + ".dae"
                #fpathNew = os.path.join (destPath, newFileName)
                #removeFile (fpathNew) #if file exists, delete it
                #exportDAE15(fpathNew,newDoc)
                
                #3DS Export
                newFileName = destFileName + ".3ds"
                fpathNew = os.path.join (destPath, newFileName)
                removeFile (fpathNew) #if file exists, delete it
                export3DS(fpathNew,newDoc)
                
                #VRML Export
                #newFileName = destFileName + ".wrl"
                #fpathNew = os.path.join (destPath, newFileName)
                #removeFile (fpathNew) #if file exists, delete it
                #exportVRML2(fpathNew,newDoc)
                
                nFiles += 1

                c4d.documents.KillDocument(newDoc)
           
    print ("                ")    
    print ("*** GOOD JOB ***")   
    print ("Folders checked:  " + str(nFolders))
    print ("Files converted:  " + str(nFiles))
    print ("End of conversion")
     
    c4d.EventAdd()


if __name__=='__main__':
    main()