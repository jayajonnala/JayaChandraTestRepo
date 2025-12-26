
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Distribution process to DS stores - SW41 (Fresh meat) - p3
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
'.................Test Script Name :Test_Distribution process to DS stores - SW41 (Fresh meat) - p3
'.................Author : TCS 	   :Raushan
'................ Creation Date    :16th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Distribution process to DS stores - SW41 (Fresh meat) - p3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Distribution process to DS stores - SW41 (Fresh meat) - p1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode VL03N----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

'Enter the Outbound Delivery No
Call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_VL03N_4004_OUTBOUND_DELIVERY,FALSE)
Call TakeScreenShot()
Call PressEnter()
Wait(3)
Call TakeScreenShot()

'Click on Document Flow
Call ClickButton("Document Flow   \(F7\)",False)
Wait(2)
Call TakeScreenShot()

'Click on GUI Tree
' ActivateNodeGuiTree(treeIndex, itemPath)
'Call ActivateNodeGuiTree(0,"#1;#1;#1")

' ActivateItemGuiTree(treeIndex, itemPath, itemText)
Call ActivateItemGuiTree(0,"#1;#1;#2","#2")
Wait(2)
Call TakeScreenShot()

'Get the Document No
Call GetGridContentByTitle("TF to stck in trans",0,"Doc.no.",1,"DT_DOCUMENT_NUM_OUTPUT")
'
'----------------------Tcode MIGO----------------------------

Call SetTcode(DT_VL03N_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_VL03N_0100_OKCD)

'------------------------------Display Article Document Details-------------------------------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetCombo("GODYNPRO-ACTION","Display")
'Call SetCombo("GODYNPRO-REFDOC","Article Document")
Call SetComboByKey("GODYNPRO-REFDOC","R02")
Call SetTextboxNoLabel("GODYNPRO-MAT_DOC",0,DT_VL03N_2010_GODYNPROMAT_DOC,False)
Call SetTextboxNoLabel("GODYNPRO-DOC_YEAR",0,DT_VL03N_2010_GODYNPRODOC_YEAR,False)
Call PressEnter() 
Wait(2)
Call TakeScreenShot()

Call ClickButtonIfExist("Open detail data",False)
'Navigate to messages Tab
Call SelectTab("TS_GOITEM","Messages",False)
Wait(1)
Call TakeScreenShot()

'Click on Display Item
Call ClickButtonIfExist("Display outputs",False)
Wait(2)
Call TakeScreenShot()

'Verify the Output Type
Call VerifyTableCellContent(1,"Output Type","SAPDV70ATC_NAST3",DT_VL03N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_0)

'Click on Back
Call ClickButtonIfExist("Back   \(F3\)",False)
Wait(2)

'Navigate to the Document Details
Call SelectTab("TS_GOHEAD","Doc. info",False)
Wait(1)
Call TakeScreenShot()

Call ClickButton("FI Documents",False)
Wait(5)
Call TakeScreenShot()

'Get document No
Call GetGridContentByTitle("Documents in Accounting",0,"Doc. Number",1,"DT_DOC_NO_ACCOUNTING")

'Click on Cancel
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Wait(2)
Call ClickButtonIfExist("Close Detail Data",False)

'Validate the Movement Type
Call VerifyTableCellContent(1,"Movement type","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_0)
Call VerifyTableCellContent(2,"Movement type","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_1)
Call VerifyTableCellContent(3,"Movement type","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_2)
Call VerifyTableCellContent(4,"Movement type","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_3)
Call VerifyTableCellContent(5,"Movement type","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_4)
Call VerifyTableCellContent(6,"Movement type","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_5)
Call VerifyTableCellContent(7,"Movement type","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_6)
Call VerifyTableCellContent(8,"Movement type","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_7)
Call VerifyTableCellContent(9,"Movement type","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_8)
Call VerifyTableCellContent(10,"Movement type","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_9)
Call VerifyTableCellContent(11,"Movement type","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_10)
Call VerifyTableCellContent(12,"Movement type","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_11)
Call VerifyTableCellContent(13,"Movement type","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_12)
Call VerifyTableCellContent(14,"Movement type","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_13)
Call VerifyTableCellContent(15,"Movement type","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_14)
Call VerifyTableCellContent(16,"Movement type","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_15)

'Validate the Direction
Call VerifyTableCellContent(1,"Direction","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_0)
Call VerifyTableCellContent(2,"Direction","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_1)
Call VerifyTableCellContent(3,"Direction","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_2)
Call VerifyTableCellContent(4,"Direction","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_3)
Call VerifyTableCellContent(5,"Direction","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_4)
Call VerifyTableCellContent(6,"Direction","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_5)
Call VerifyTableCellContent(7,"Direction","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_6)
Call VerifyTableCellContent(8,"Direction","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_7)
Call VerifyTableCellContent(9,"Direction","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_8)
Call VerifyTableCellContent(10,"Direction","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_9)
Call VerifyTableCellContent(11,"Direction","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_10)
Call VerifyTableCellContent(12,"Direction","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_11)
Call VerifyTableCellContent(13,"Direction","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_12)
Call VerifyTableCellContent(14,"Direction","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_13)
Call VerifyTableCellContent(15,"Direction","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_14)
Call VerifyTableCellContent(16,"Direction","SAPLMIGOTV_GOITEM",DT_VL03N_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_15)


'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

