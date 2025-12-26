
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Purchasing Meat with collective PO for standard article p4
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

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Purchasing Meat with collective PO for standard article p4
'.................Author : TCS 	   :Raushan
'................ Creation Date    :16th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Purchasing Meat with collective PO for standard article p4"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\DS\RETAIL\DT_Purchasing Meat with collective PO for standard article p4_TASE.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'------------------------------------------MIGO------------------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

''------------------------------Display Article Document Details-------------------------------------------------

Call SetCombo("GODYNPRO-ACTION","Goods Receipt")
Call SetCombo("GODYNPRO-REFDOC","Purchase Order")
Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER",0,DT_PO_NUMBER,False)
Call TakeScreenShot()
Call PressEnter()   
Call SetTextBox("Delivery Note","GOHEAD-LFSNR",0,DT_MIGO_0110_DELIVERY_NOTE,False)
Call PressEnter() 
Call TakeScreenShot()

Call SelectCheckbox("GODYNPRO-DETAIL_TAKE", 0, "ON", False)
Call TakeScreenShot()

Call SelectTab("TS_GOITEM","Batch",False)
Call TakeScreenShot()

Call SetGridData("",1,"Documentary Batch - Batch No.",DT_MIGO_0201_GRIDCELL_0_DB_NO_OCC2,False)
Call SetGridData("",1,"Qty in unit of entry",DT_MIGO_0201_GRIDCELL_0_QTY_IN_UN_OF_ENTRY_OCC4,False)
Call TakeScreenShot()
Call PressEnter() 

Call ClickButton("Next Item",False)
Call TakeScreenShot()
Call ClickButton("Check Entries   \(F7\)",False)
Call TakeScreenShot()
Call ClickButton("Back   \(F3\)",False)

''''Click on save
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot()
WAIT 5
Call ClickButton("Back   \(F3\)",False)
Call GetStatusBar("item1","DT_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Article Document "&DT_STATUSBAR_OUTPUT& " posted")

Call LogOff()
Call FinalStatus ()


'
''*********************************************End Of Script*********************************************************************
'
''Navigate to the Document Details
'Call SelectTab("TS_GOHEAD","Doc. info",False)
'Wait(1)
'Call TakeScreenShot()
'
''Click on F1 Document
'Call ClickButton("FI Documents",False)
'Wait(2)
'Call TakeScreenShot()
'
''Get The Document No
'Call GetGridContentByTitle("Documents in Accounting",0,"Doc. Number",1,DT_ACCT_DOCUMENTNO)
'
''Click on Cancel Button
'Call ClickButton("Cancel   \(F12\)",True)
'Wait(1)
'
'Call ClickButtonIfExist("Close Detail Data",False)
'Call ClickButtonIfExist("Close Header Data",False)
'
'
'
''Validate the Movement Type
'Call VerifyTableCellContent(1,"Movement type","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_0)
'Call VerifyTableCellContent(2,"Movement type","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_1)
'Call VerifyTableCellContent(3,"Movement type","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_2)
'Call VerifyTableCellContent(4,"Movement type","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_3)
'Call VerifyTableCellContent(5,"Movement type","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_4)
'Call VerifyTableCellContent(6,"Movement type","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_5)
'Call VerifyTableCellContent(7,"Movement type","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_6)
'Call VerifyTableCellContent(8,"Movement type","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_7)
'Call VerifyTableCellContent(9,"Movement type","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_8)
'Call VerifyTableCellContent(10,"Movement type","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_9)
'Call VerifyTableCellContent(11,"Movement type","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_10)
'Call VerifyTableCellContent(12,"Movement type","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_11)
'
''Validate the Direction
'Call VerifyTableCellContent(1,"Direction","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_0)
'Call VerifyTableCellContent(2,"Direction","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_1)
'Call VerifyTableCellContent(3,"Direction","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_2)
'Call VerifyTableCellContent(4,"Direction","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_3)
'Call VerifyTableCellContent(5,"Direction","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_4)
'Call VerifyTableCellContent(6,"Direction","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_5)
'Call VerifyTableCellContent(7,"Direction","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_6)
'Call VerifyTableCellContent(8,"Direction","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_7)
'Call VerifyTableCellContent(9,"Direction","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_8)
'Call VerifyTableCellContent(10,"Direction","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_9)
'Call VerifyTableCellContent(11,"Direction","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_10)
'Call VerifyTableCellContent(12,"Direction","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_11)
'
'Call ClickButtonIfExist("Open header data",False)
'Call ClickButtonIfExist("Open detail data",False)
'
''------------------------'Log Off From Applicaton-------------------------

