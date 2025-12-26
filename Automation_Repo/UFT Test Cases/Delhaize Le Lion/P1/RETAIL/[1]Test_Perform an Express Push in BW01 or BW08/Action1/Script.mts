
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Perform an Express Push in BW01 or BW08
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	DataRowSet= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrTestCaseName = "Test_Perform an Express Push in BW01 or BW08"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\DLL\RETAIL\DT_Perform a Zero Price Push Order in BW01 or BW08_TASE.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call TakeScreenShot()
''
'--------------------------------------------  MIGO ----------------------------------------------

Call SetComboByKey("GODYNPRO-ACTION",DT_MIGO_0010_GODYNPROACTION)
Call SetComboByKey("GODYNPRO-REFDOC",DT_MIGO_0010_GODYNPROREFDOC)
Call SetTextboxNoLabel("GODEFAULT_TV-BWART",0,DT_MIGO_0010_GODEFAULT_TVBWART,False)
Call PressEnter()
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot()

Call SelectTab("TS_GOITEM","Article",False)
Call SetTextbox("Article","GOITEM-MAKTX",0,DT_MIGO_0310_ARTICLE,False)
Call PressEnter() 
Call TakeScreenShot()

Call SelectTab("TS_GOITEM","Quantity",False)
Call TakeScreenShot()
Call SetTextbox("Qty in Unit of Entry","GOITEM-ERFMG",0,DT_MIGO_0315_QTY_IN_UNIT_OF_ENTRY,False)
Call TakeScreenShot()
Call PressEnter() 
Call TakeScreenShot()

Call SelectTab("TS_GOITEM","Where",False)
Call TakeScreenShot()
Call SetTextbox("Site","GOITEM-NAME1",0,DT_MIGO_0325_SITE,False)
Call SetTextbox("Storage Location","GOITEM-LGOBE",0,DT_MIGO_0325_STORAGE_LOCATION,False)
Call TakeScreenShot()
Call PressEnter() 
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot()

Call GetStatusBar("item1","DT_ART_DOC_NUMBER_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_ART_DOC_NUMBER_OUTPUT",DT_ART_DOC_NUMBER)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(lcase(DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR_OCC1))

Call SetTcode(DT_MIGO_0001_OKCD) 
Call PressEnter()     ' 
Call TakeScreenShot()
''
'--------------------------------------------  WA01 ----------------------------------------------

Call SetTextbox("Allocation Table Type","AUKO-AUFAR",0,DT_MIGO_0100_ALLOCATION_TABLE_TYPE,False)
Call SetTextbox("Purchasing Organization","AUKO-EKORG",0,DT_MIGO_0100_PURCHASING_ORGANIZATION,False)
Call SetTextbox("Purchasing Group","AUKO-EKGRP",0,DT_MIGO_0100_PURCHASING_GROUP,False)
Call TakeScreenShot()
Call PressEnter() 

Call SetTextbox("Site Delivery Date","RM01A-WEFDT",0,ConvertDate(DT_MIGO_0122_SITE_DELIVERY_DATE),False)
Call SetTextbox("Allocation Table","AUKO-BEZCH",0,DT_MIGO_0122_ALLOCATION_TABLE,False)
Call TakeScreenShot()
Call PressEnter() 
Call TakeScreenShot()

Call SelectMenuBar("Allocation table;Create with Reference;User Exit Article Selection")
Call SetTextbox("Directory","DY_PATH",0,DT_MIGO_0200_DIRECTORY,True)
Call SetTextbox("File Name","DY_FILENAME",0,DT_MIGO_0200_FILE_NAME,True)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)
Call PressEnter()

Wait(5)
Call TakeScreenShot()
Call SelectRowGuiTable("SAPML01ATC_0122","Article",DT_MIGO_0310_ARTICLE,False)
Call ClickButton("Fast item change   \(Ctrl\+F9\)",False)

Call SelectRadioButton("G_DATEN_ALLE_POS","Also overwrite existing values",True)
Call SetTextbox("Recipient Determination","AUPO-NREMFIN",0,DT_MIGO_0305_RECIPIENT_DETERMINATION,True)
Call SetTextbox("Allocation Strategy","AUPO-ASTRA",0,DT_MIGO_0305_ALLOCATION_STRATEGY,True)
Call SetTextbox("Item category","AUPO-APSTP",0,DT_MIGO_0305_ITEM_CATEGORY,True)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",True)
Call VerifyTextBoxContent("Information Message","MESSTXT1",0,lcase(DT_MIGO_0010_CHECK_TEXT_OF_MESSTXT1),True)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButton("Save",True)
Call TakeScreenShot()

Call GetStatusBar("item1","DT_ALLOCATION_TABLE_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_ALLOCATION_TABLE_OUTPUT",DT_ALLOCATION_TABLE)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(lcase(DT_MIGO_0100_CHECK_TEXT_OF_STATUSBAR_OCC1))

Call LogOff()
Call FinalStatus ()



