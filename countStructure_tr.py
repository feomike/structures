# ---------------------------------------------------------------------------
# countStructure_tr.py
# Created: Nov. 2007
#   Mike Byrne
#   Office of Statewide Health Planning and Development
#   this script counts the number of generated structures per census Tract
#   i do this so we can check the variation generated against Claritas
# ---------------------------------------------------------------------------

# Import system modules
from win32com.client import Dispatch
import sys, os, random, string, array

# Create the Geoprocessor object
gp = Dispatch("esriGeoprocessing.GpDispatch.1")

#path names can use "/" or "\\"
theStru = "D:/Structures/PostProcessing/Regions/"
theRegionFC = "D:/Broadband/BB_GIS_Lib.mdb/Regions"
theCTYFC = "D:/Broadband/BB_GIS_Lib.mdb/County"
theTRFC = "D:/Broadband/BB_GIS_Lib.mdb/Census_Tracts"
outFileAdd = "D:/Structures/stru_tr_count.txt"   #output File

#make sure you can handle the errors
try:
    gp.AddMessage("starting the process...")
    myFileAdd = open(outFileAdd, 'w')
    myFileAdd.write("CENSUS_KEY,NumStru" + "\n")
    Regions = ["CC", "BA","ES","IE","LO","ML","NC","NS","SB","SJ","SM","SV"] 
    for myRegion in Regions:
        theStruLyr = "StructureLyr" + myRegion        
        theStruFC = theStru + "Region_" + myRegion + ".shp"
        gp.MakeFeatureLayer_management(theStruFC, theStruLyr)
        gp.AddMessage(myRegion)
        theRegionLyr = "Region_" + myRegion
        theRegionQry = "[Regs] = '" + myRegion + "'"
        gp.MakeFeatureLayer_management(theRegionFC, theRegionLyr, theRegionQry)
        #get the counties w/i this
        theCtyLyr = "CountyLayer" + myRegion
        gp.MakeFeatureLayer_management(theCTYFC, theCtyLyr)
        gp.SelectLayerByLocation_management(theCtyLyr, "HAVE_THEIR_CENTER_IN", theRegionLyr, "", "NEW_SELECTION")
        #get the Census Tracts w/i this county for loop driving
        theTRDriveLyr = "TractLyr" + myRegion
        gp.MakeFeatureLayer_management(theTRFC, theTRDriveLyr)
        gp.SelectLayerByLocation_management(theTRDriveLyr, "HAVE_THEIR_CENTER_IN", theCtyLyr, "", "NEW_SELECTION")
        trCur = gp.SearchCursor(theTRDriveLyr)
        trRow = trCur.Next()
        while trRow <> None:
            #get a census key
            myKey = trRow.GetValue("Census_Key")
            theTRQry = "[Census_key] = '" + myKey + "'"
            theTractLyr = "Tract" + myKey
            gp.MakeFeatureLayer_management(theTRFC, theTractLyr, theTRQry)
            gp.SelectLayerByLocation_management(theStruLyr, "INTERSECT", theTractLyr, "", "NEW_SELECTION")
            #get the count
            theVal = gp.GetCount_management(theStruLyr)
            #write out the
            gp.AddMessage("Writing out Census_Key: " + myKey + " with " + str(theVal) + " Structures")
            myFileAdd.write(myKey + "," + str(theVal) + "\n")
            gp.Delete_management(theTractLyr)
            del theVal, myKey, theTractLyr
            trRow = trCur.Next()
        gp.Delete_management(theTRDriveLyr)
        gp.Delete_management(theRegionLyr)
        gp.Delete_management(theCtyLyr)
        gp.Delete_management(theStruLyr)
        del trRow, trCur, theTRDriveLyr, theRegionLyr, theCtyLyr, theStruLyr, theStruFC, theTRQry
    #clean things up; delete variables
    myFileAdd.close()
    del myRegion, Regions, outFileAdd, theTRFC, theCTYFC, theRegionFC, theStru
    gp.AddMessage("You rock, structures counted inside tracts")
except:
    gp.AddMessage(gp.GetMessage(0))
    gp.AddMessage(gp.GetMessage(1))
    gp.AddMessage(gp.GetMessage(2))
    gp.AddMessage("Something bad happend...")
    
