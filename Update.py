import arcpy #importing arcpy library in arcgis
import xlrd # for only read the excelsheet
import xlwt # for writing in the excelsheet
path=r'D:\Excel\'abc.xls' #path for reading file
wb=xlrd.open_workbook(path) #open workbook
sheet=wb.sheet_by_name('Sheet3') # open worksheet by name
nwb=xlwt.Workbook() #intialise excelsheet for writing mode
nsheet=nwb.add_sheet("new",cell_overwrite_ok=True) # give sheet name 
def searching(): #function for selecting attrbutes and updating values
     roe=0
     rr=300
     for x in range(1,rr):
         t=str(sheet.cell(x,2).value) #for first condition attrbute value
         v=str(sheet.cell(x,1).value) #for second condition attrbute value 
         g=int(sheet.cell(x,3).value) #for third condition attrbute value
         n=str(sheet.cell(x,0).value) #for updating attrbute value
         #arcpy function for selecting value from given layer
         arcpy.SelectLayerByAttribute_management("Layer","NEW_SELECTION","first_column'='{}' and 'second_column'='{}' and 'third_column'='{}'".format(t,v,g))
         #"NEW_SELECTION" selection type of attribute for more read "arcpy.SelectLayerByAttribute_management" in ArcGis help desk
         print('Counting=',x,'Remaining=',rr-x) #for understanding the running code
         cursor=arcpy.UpdateCursor("Layer") #update function in arcpy for the above condition
         for xy in cursor:
             xy.setValue('New',n) # 'New' is column name which we have to update
             cursor.updateRow(xy) # its update the column in Layer
         cursorr=arcpy.SearchCursor("Layer")
         for wid in cursorr: # writing into the excel sheet with iterating columns and rows
             value=wid.getValue('New')
             val=wid.getValue('third_column'')
             area=wid.getValue('Area_Ha')
             nsheet.write(roe,0,t)
             nsheet.write(roe,1,v)
             nsheet.write(roe,2,g)
             nsheet.write(roe,3,value)
             nsheet.write(roe,4,val)        
             nsheet.write(roe,5,area)
             roe+=1
     nwb.save('xx.xls') #saved in new excel sheet
     print('COMPLETED')
