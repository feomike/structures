# ---------------------------------------------------------------------------
# countStructure.py
# Created: Oct. 2007
#   Mike Byrne
#   Office of Statewide Health Planning and Development
#   this script counts (and calculates fields) the structures
#   based on 1) the Broadband Region, 2) the Rural Urban Definition and
#   3) the Broadband speed tier
#   it requires (1) a region feature class
#               (2) a MSSA (rural urban) feature class
#               (3) a speed tier feature class
#               (4) a structures (housing units) dataset
# ---------------------------------------------------------------------------

# Import system modules
from win32com.client import Dispatch
import sys, os, random, string, array

# Create the Geoprocessor object
gp = Dispatch("esriGeoprocessing.GpDispatch.1")

#path names can use "/" or "\\"
theStru = "D:/Structures/PostProcessing/Regions/"
outFileAdd = "D:/Structures/stru_count_att.txt"   #output File

#make sure you can handle the errors
try:
    gp.AddMessage("starting the process...")
    myFileAdd = open(outFileAdd, 'w')
    myFileAdd.write(",SpeedTier1,,SpeedTier2,,SpeedTier3,,SpeedTier4,,SpeedTier5,,Urban,Rural,Total" + "\n")
    myFileAdd.write("Region,UrbanA,RuralA,UrbanA,RuralA,UrbanA,RuralA,UrbanA,RuralA,UrbanA,RuralA" + "\n")
    Regions = ["ES",] #,"BA","ES","IE","LO","ML","NC","NS","SB","SJ","SM","SV"] #all the other ones
    for myRegion in Regions:
        theString = myRegion
        theIntStr = ","
        #gp.AddMessage(myRegion)
        theStruLyr = "StructureLyr" + myRegion        
        theStruFC = theStru + "Region_" + myRegion + ".shp"
        gp.MakeFeatureLayer_management(theStruFC, theStruLyr)
        theTotStru = gp.GetCount_management(theStruLyr)
        #get total Rural and Urban in the region
        theUrbQry = "'Urban' = 1"
        theUrbLyr = "StructureUrbanLyr" + myRegion
        gp.MakeFeatureLayer_management(theStruFC, theUrbLyr, theUrbQry)
        theUrbCnt = gp.GetCount_management(theUrbLyr)
        theRurQry = "'Urban' = 0"
        theRurLyr = "StructureRuralLyr" + myRegion
        gp.MakeFeatureLayer_management(theStruFC, theRurLyr, theRurQry)
        theRurCnt = gp.GetCount_management(theRurLyr)
        gp.Delete_management(theUrbLyr)
        gp.Delete_management(theRurLyr)
        gp.Delete_management(theStruLyr)
        gp.AddMessage("totals are:")
        gp.AddMessage("Urban is:" + str(theUrbCnt) + " and Rural is: " + str(theRurCnt))
        del theRurLyr, theRurQry, theUrbLyr, theUrbQry
        Speeds = [1,2,3,4,5]
        for mySpeed in Speeds:
            Definitions = [1, 0]
            for myDef in Definitions:
                theQry = "[Tier] " + str(mySpeed) + "[Urban] = " + str(myDef)
                theDefLyr = "StruQryLyr"
                gp.MakeFeatureLayer_management(theStruFC, theDefLyr, theQry)
                #Select structures in that tier
                theVal = gp.GetCount_management(theDefLyr)
                theIntStr = theIntStr + str(theVal) + ","
                gp.Delete_management(theDefLyr)
        theString = theString + theIntStr + str(theUrbCnt) + "," + str(theRurCnt) + "," + str(theTotStru) 
        gp.AddMessage(theString)
        myFileAdd.write(theString + "\n")
    
    #clean things up; delete variables
    del theTotStru, theRurCnt, theUrbCnt, theIntStr, theString, theStruFC, theStru
    myFileAdd.close()
    gp.AddMessage("You rock, random structures generated")
except:
    gp.AddMessage(gp.GetMessage(0))
    gp.AddMessage(gp.GetMessage(1))
    gp.AddMessage(gp.GetMessage(2))
    gp.AddMessage("Something bad happend...")
    
