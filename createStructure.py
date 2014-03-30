# ---------------------------------------------------------------------------
# createStructure.py
# Created: Sept. 2007
#   Mike Byrne
#   Office of Statewide Health Planning and Development
#   this script creates a list of randomly assigned structures within each census block group
#   based on a (1) the Census SF3 number of structures within the group
#   it requires a (1) census block group geography with the field LOGRECNO
#                 (2) a table with the LOGRECNO (unique for Census Block Group) field
#                 (3) and a field with the number of structures in each Block Group
# ---------------------------------------------------------------------------

# Import system modules
from win32com.client import Dispatch
import sys, os, random, string, array

# Create the Geoprocessor object
gp = Dispatch("esriGeoprocessing.GpDispatch.1")

#ctyNum = '015'
ctyNum = sys.argv[1]
TAPointsLogical = sys.argv[2]

#path names can use "/" or "\\"
bgFC = "D:/Structures/Structure_BG_Processing.mdb/bg_" + ctyNum #"D:\\Structures\\Data.mdb\\BlockGroup"   #Block Group Feature Class
bgFCField = "LOGRECNO"                              #unique Block Group Field; default is Logrecno
roadFC = "D:/Structures/Structure_Road_Processing.mdb/Street_" + ctyNum #"D:\\Structures\\Data.mdb\\Streets"    #street feature class
stTable = "D:/Structures/Data.mdb/cen00yrb03_1"            #stru_" + ctyNum  #Structure Table
stbgField = "LOGRECNO"                           #Structure Block Group Field; default is Logrecno; a relate field connecting to the bgFC
stField = "H034001"                                 #Structure field; default is H034001
#TAPointFC = "Database Connections/Connection to egis-data-editor.sde/Vector.EGIS.Transportation/VECTOR.EGIS.Address_Points"     #"C:\\GIS-Data\\StructureProcessing.mdb\\TA_Points"
TAPointFC = "D:/Address_Points_037.shp" #laptop connection
outFileAdd = "D:/Structures/stru_" + ctyNum + "Address.txt"   #output File
outFileEvn = "D:/Structures/stru_" + ctyNum + "Event.txt"   #output File

#make sure you can handle the errors
try:
    myFileAdd = open(outFileAdd, 'w')
    myFileAdd.write(bgFCField + "`SUMELEV_090`LOID`Address`ZIPCode" + "\n")
    myFileEvn = open(outFileEvn, 'w')
    myFileEvn.write(bgFCField + ",SUMELEV_090,LOID,X,Y" + "\n")
    gp.AddMessage("starting the process...")
    gp.OverWriteOutput = 1
    #make feature layers for the road feature class
    rdStr = "[FCC] <> 'A11' AND [FCC] <> 'A12' AND [FCC] <> 'A15' AND [FCC] <> 'A16'  " #remove all interstates
    rdStr = rdStr + "AND [FCC] <> 'A17' AND [FCC] <> 'A18' AND [FCC] <> 'A60' AND [FCC] <> "
    rdStr = rdStr + "'A63' AND [FCC] <> 'A64' AND [FCC] <> 'A65' AND [FCC] <> 'A66' AND [FCC] <> 'A69'"
    roadLyr = "StreetLayer" + ctyNum
    gp.MakeFeatureLayer_management(roadFC, roadLyr, rdStr)
    stView = "StructureView"
    myWhere = "[" + stField + "] > 0 AND [SUMLEV_090] LIKE '06" + ctyNum + "*'"
    gp.MakeTableView_management (stTable, stView, myWhere)
    #set up a cursor loop for obtaining each Census Block record and structure value
    stCur = gp.UpdateCursor(stView)
    stRow = stCur.Next()
    
    while stRow <> None:
        #get Census Block Unique ID and number of Structures from the structures table
        logRecNo = stRow.GetValue(stbgField)
        cblFIPS = stRow.GetValue("SUMLEV_090")
        numStructures = stRow.GetValue(stField)
        #make a Feature Layer of just the single Census Block
        stQyr = "[" + bgFCField + "] = '" + logRecNo + "'"
        bgLyr = "BlockLayer" + logRecNo
        #make a layer of just that census bloc
        gp.MakeFeatureLayer_management(bgFC, bgLyr, stQyr)
        TAPointCnt = 0
        if TAPointsLogical == "Yes":
            #retrieve the number of TA_Points in the Block, so you can subtract those from the x-loop
            TAPointLyr = "TAPointLayer"
            gp.MakeFeatureLayer_management(TAPointFC,TAPointLyr)
            #select TAPoints w/in Selected
            gp.SelectLayerByLocation_management(TAPointLyr, "INTERSECT", bgLyr)
            TAPointCnt = gp.GetCount_management(TAPointLyr)
            gp.Delete_management(TAPointLyr)
            del TAPointLyr
        else:
            TAPointCnt = 0
        numStructures = numStructures - TAPointCnt
        
        #select streets within bgLayer, and load into an array
        gp.SelectLayerByLocation_management(roadLyr, "INTERSECT", bgLyr)
        myArray = array.array('L')
        cnt = gp.GetCount_management(roadLyr)
        roadCur = gp.SearchCursor(roadLyr)
        roadRow = roadCur.Next()  # move to that record number
        while roadRow:
            myArray.append(roadRow.GetValue("ObjectID"))
            #gp.AddMessage(roadRow.GetValue("ObjectID"))
            roadRow = roadCur.Next()
        myArray.tolist()
        #clean up
        gp.Delete_management(bgLyr)
        del TAPointCnt, bgLyr, stQyr, roadRow, roadCur
        #loop through for each number of structures
        x = 1
        while x <= numStructures and gp.GetCount_management(roadLyr) > 0:
            ##select a random lineSegment within the selected roads array above
            myLine = random.randint(0, (cnt - 1))
            ##Make a random Selection between 0 or 1
            mySide = random.randint(0, 1) #0 = Left, 1 = Right
            if mySide == 0:
                theZone = "ZIPL"
                theFromAdd = "L_F_ADD"
                theToAdd = "L_T_ADD"
            else:
                theZone = "ZIPR"
                theFromAdd = "R_F_ADD"
                theToAdd = "R_T_ADD"
            #move to record in the array
            myOID = myArray[myLine]
            myRoadLyr = "road" + str(x)
            myRdStr = "[ObjectID] = " + str(myOID)
            gp.MakeFeatureLayer_management(roadFC, myRoadLyr, myRdStr)
            roadCur = gp.SearchCursor(myRoadLyr)
            roadRow = roadCur.Next()  # move to that record number
            #get an address
            myFromAdd = string.strip(roadRow.GetValue(theFromAdd))
            myToAdd = string.strip(roadRow.GetValue(theToAdd))
            if myFromAdd.isdigit() == 1 and myToAdd.isdigit() == 1:
                #gp.AddMessage("Finding a random address ...")
                if int(myFromAdd) < int(myToAdd):
                    myHouse = random.randint(int(myFromAdd), int(myToAdd))
                else:
                    myHouse = random.randint(int(myToAdd), int(myFromAdd))
                #get a Prefix, StreetName, Type, and ZIPCode
                myPre = string.strip(str(roadRow.GetValue("PREFIX")))
                if len(myPre) == " ":
                    myPre = ""
                myName = string.strip(str(roadRow.GetValue("NAME")))
                myType = string.strip(str(roadRow.GetValue("TYPE")))
                if myType == " ":
                    myType = ""
                mySuffix = string.strip(str(roadRow.GetValue("SUFFIX")))
                if mySuffix == " ":
                    mySuffix = ""
                myZone = string.strip(str(roadRow.GetValue(theZone)))
                ##write out the record to a file
                #
                myStreet = string.strip(str(myPre + " " + myName + " " + myType)) + " " + mySuffix
                myStreet = str(myHouse) + " " + myStreet +  "`" + myZone
                gp.AddMessage(logRecNo + "`" + myStreet + " - for; " + str(x) + " of " + str(numStructures) + " structure")
                myFileAdd.write(logRecNo + "`" + cblFIPS + "`" + str(myOID) + "`" + myStreet + "\n")
                del myHouse, myPre, myName, myType, myZone, mySuffix, myStreet
            else:  #this is where you select a random segment of the line, and get an x,y for that segment
                #gp.AddMessage("Finding a random line segment ...")
                feat = roadRow.shape
                myPart = 0
                while myPart < feat.PartCount: #ensuring it is not multi part
                    roadArray = feat.GetPart(myPart)
                    roadArray.Reset
                    myVert = random.randint(1, roadArray.Count)
                    myFeat = 0
                    while myFeat < myVert: #loop through each part until you get to the random number
                        pnt = roadArray.Next()
                        myFeat = myFeat + 1
                    myX = str(pnt.x)
                    myY = str(pnt.y)
                    myPart = myPart + 1
                gp.AddMessage(logRecNo + "," + myX + "," + myY + " - for; " + str(x) + " of " + str(numStructures) + " structure")
                myFileEvn.write(logRecNo + "," + cblFIPS + "," + str(myOID) + "," + myX + "," + myY + "\n")
                #cleanup
                del myX, myY, pnt, myFeat, myVert, roadArray, myPart, feat
            del myLine, mySide, theZone, theFromAdd, theToAdd, myFromAdd, myToAdd
            del myOID, roadRow, roadCur
            x = x + 1   #last line of number of structures loop
        #clean up things
        del x, numStructures, logRecNo, cblFIPS, cnt, myArray
        stRow.SetValue(stField,0)
        stCur.UpdateRow(stRow)
        stRow = stCur.Next()
    
    #clean things up; delete variables
    myFileAdd.close()
    myFileEvn.close()
    gp.Delete_management(stView)
    gp.Delete_management(roadLyr)
    del stRow, stCur, myWhere, stView, roadLyr, myFileAdd, myFileEvn
    del ctyNum, outFileAdd, outFileEvn, stTable, stbgField, stField
    del bgFCField, TAPointFC, roadFC, bgFC
    gp.AddMessage("You rock, random structures generated")
except:
    gp.AddMessage(gp.GetMessage(0))
    gp.AddMessage(gp.GetMessage(1))
    gp.AddMessage(gp.GetMessage(2))
    gp.AddMessage("Something bad happend...")
    
