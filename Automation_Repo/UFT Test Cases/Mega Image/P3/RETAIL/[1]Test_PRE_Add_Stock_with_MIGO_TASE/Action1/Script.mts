'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	GetRowNo= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'reload DS to update dates and calculations
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_PRE_Add_Stock_with_MIGO_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 4th June
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName="Test_PRE_Add_Stock_with_MIGO_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_PRE_Add_Stock_with_MIGO_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)
''''Login'''
'DataRowSet=2
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
'''----------------------Tcode MIGO----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetComboByKey("GODYNPRO-ACTION",DT_MIGO_0010_GODYNPROACTION)
'Capture the screenshot
Call TakeScreenShot()

Call SetComboByKey("GODYNPRO-REFDOC",DT_MIGO_0010_GODYNPROREFDOC)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()

'
Call SelectTab("TS_GOITEM",DT_MIGO_0300_QUANTITY,False)
Call SetTextbox("Qty in Unit of Entry","GOITEM-ERFMG","",DT_MIGO_0315_MOVEMENT_TYPE,False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextboxNoLabel("GODEFAULT_TV-BWART",0,DT_MIGO_0010_GODEFAULT_TVBWART,False)
Call TakeScreenShot()
Call SetTextbox("Article Slip","GOHEAD-MTSNR","",DT_MIGO_0112_ARTICLE_SLIP,False)
Call TakeScreenShot()

Call SelectTab("TS_GOITEM","Article",False)
Call SetTextbox("Article","GOITEM-MAKTX","",DT_MIGO_0310_MOVEMENT_TYPE,False)
Call SelectTab("TS_GOITEM",DT_MIGO_0300_QUANTITY,False)
Call SetTextbox("Qty in Unit of Entry","GOITEM-ERFMG","",DT_MIGO_0315_MOVEMENT_TYPE,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TS_GOITEM",DT_MIGO_0300_WHERE,False)
Call SetTextbox("Movement type","GOITEM-BWART","",DT_MIGO_0325_MOVEMENT_TYPE,False)
Call SetTextbox("Site","GOITEM-NAME1","",DT_MIGO_0325_SITE,False)
Call SetTextbox("Storage Location","GOITEM-LGOBE","",DT_MIGO_0325_STORAGE_LOCATION,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()
Call SelectTab("TS_GOITEM",DT_MIGO_0300_PARTNER,False)
Call SetTextbox("Vendor","GOITEM-VENDORNAME","",DT_MIGO_0340_VENDOR,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

'Post the Article No
Call ClickButton("Post   \(Ctrl\+S\)",False)
Wait(2)
Call ClickButtonIfExist("Save",True) 
Wait(2)
'Capture the screenshot
Call TakeScreenShot()
'save it to data sheet
Call GetStatusBar("item1","DT_GR_NUMBER_OUTPUT")
'reload data sheet
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'verify statusbar with datasheet feed
Call VerifyStatusBar(DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR_OCC1)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()


