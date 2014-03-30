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
theStru = "D:/Structures/PostProcessing/Events/Region_Event_Matched.mdb/"
theStru = "D:/Structures/PostProcessing/Regions/"
theRegionFC = "D:/Broadband/BB_GIS_Lib.mdb/Regions"
theMSSAFC = "D:/Broadband/BB_GIS_Lib.mdb/MSSA2005"
theBBTierFC = "D:/Broadband/BB_GIS_Lib.mdb/CBTF_Poly"
outFileAdd = "D:/Structures/stru_count.txt"   #output File

#make sure you can handle the errors
try:
    gp.AddMessage("starting the process...")
    myFileAdd = open(outFileAdd, 'w')
    myFileAdd.write(",SpeedTier1,,SpeedTier2,,SpeedTier3,,SpeedTier4,,SpeedTier5,,Urban,Rural,Total" + "\n")
    myFileAdd.write("Region,UrbanA,RuralA,UrbanA,RuralA,UrbanA,RuralA,UrbanA,RuralA,UrbanA,RuralA" + "\n")
    Regions = ["SM", "SV"] #,"BA","ES","IE","LO","ML","NC","NS","SB","SJ","SM","SV"] #all the other ones
    for myRegion in Regions:
        theString = myRegion
        theIntStr = ","
        #gp.AddMessage(myRegion)
        theRegionLyr = "Region_" + myRegion
        theRegionQry = "[Reg] = '" + myRegion + "'"
        gp.MakeFeatureLayer_management(theRegionFC, theRegionLyr, theRegionQry)
        theStruLyr = "StructureLyr" + myRegion        
        theStruFC = theStru + "Region_" + myRegion + ".shp"
        gp.CalculateField_management(theStruFC,"Tier","0")
        gp.CalculateField_management(theStruFC,"Urban","0")
        gp.MakeFeatureLayer_management(theStruFC, theStruLyr)
        theTotStru = gp.GetCount_management(theStruLyr)
        Speeds = [1,2,3,4,5]
        for mySpeed in Speeds:
            theSpdQry = "[Grid_Code] = " + str(mySpeed)
            theSpdLyr = "SpeedLyr" + str(mySpeed)
            gp.MakeFeatureLayer_management(theBBTierFC, theSpdLyr, theSpdQry)
            Definitions = ["=", "<>"]
            for myDef in Definitions:
                theDefQry = "[Definition] " + myDef +  "'Urban'"
                theDefLyr = "DefinitionLyr"
                gp.MakeFeatureLayer_management(theMSSAFC, theDefLyr, theDefQry)
                #Select structures in that tier
                gp.SelectLayerByLocation_management(theStruLyr, "INTERSECT", theDefLyr, "", "NEW_SELECTION")
                if myDef == "=":
                    gp.CalculateField_management(theStruLyr,"Urban",1)
                    theUr = gp.GetCount_management(theStruLyr)
                gp.SelectLayerByLocation_management(theStruLyr, "INTERSECT", theSpdLyr, "", "SUBSET_SELECTION")
                gp.CalculateField_management(theStruLyr,"Tier",mySpeed)
                theVal = gp.GetCount_management(theStruLyr)
                theIntStr = theIntStr + str(theVal) + ","
                gp.Delete_management(theDefLyr)
            gp.Delete_management(theSpdLyr)
        gp.Delete_management(theStruLyr)
        theString = theString + theIntStr + str(theUr) + ",," + str(theTotStru) 
        gp.AddMessage(theString)
        myFileAdd.write(theString + "\n")
    
    #clean things up; delete variables
    myFileAdd.close()
    gp.AddMessage("You rock, random structures generated")
except:
    gp.AddMessage(gp.GetMessage(0))
    gp.AddMessage(gp.GetMessage(1))
    gp.AddMessage(gp.GetMessage(2))
    gp.AddMessage("Something bad happend...")
    
