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

#path names can use "/" or "\\"
ctyFC = "Database Connections/Connection to egis-data.sde/Vector.EGIS.AdministrativeUnits/Vector.EGIS.Counties"
TAPointFC = "Database Connections/Connection to egis-data.sde/Vector.EGIS.Transportation/VECTOR.EGIS.Address_Points"    
outFile = "D:/Structures/cty_count.txt"   #output File

#make sure you can handle the errors
try:
    myFile = open(outFile, 'w')
    myFile.write(bgFCField + "FIPS,TA_Count" + "\n")
    gp.AddMessage("starting the process...")
    gp.OverWriteOutput = 1
    TALyr = "TALyr"
    gp.MakeFeatureLayer_management(TAPointFC, TALyr)
    #make feature layers for the road feature class
    num = 1
    while num < 59:
        #get Census Block Unique ID and number of Structures from the structures table
        ctyLyr = "CountyLayer" + str(num)
        ctyQry = "[OBJECTID] = " + str(num)
        gp.MakeFeatureLayer_management(ctyFC, ctyLyr, ctyQry)
        ctyCur = gp.SearchCursor(ctyLyr)
        ctyRow = ctyCur.Next()  # move to that record number
        FIPS = ctyRow.GetValue("CNTY_FIPS")
        del ctyRow, ctyCur
        gp.SelectLayerByLocation_management(TALyr, "INTERSECT", ctyLyr)
        myFile.write(FIPS + "," + gp.GetCount_management(TALyr) + "\n")
        gp.Delete_management(ctyLyr)
        del FIPS, ctyQry, ctyLyr
        num = num + 1

    
    #clean things up; delete variables
    myFile.close()
    myFileEvn.close()
    gp.Delete_management(TALyr)
    del num, TALyr, 
    gp.AddMessage("You rock, random structures generated")
except:
    gp.AddMessage(gp.GetMessage(0))
    gp.AddMessage(gp.GetMessage(1))
    gp.AddMessage(gp.GetMessage(2))
    gp.AddMessage("Something bad happend...")
    
