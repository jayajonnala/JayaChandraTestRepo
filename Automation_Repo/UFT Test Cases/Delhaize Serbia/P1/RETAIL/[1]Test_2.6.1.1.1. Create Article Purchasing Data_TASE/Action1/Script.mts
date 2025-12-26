

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_2.6.1.1.1. Create Article Purchasing Data
'.................Author : TCS 	   :
'................ Creation Date    : 
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

gstrTestCaseName = "Test_2.6.1.1.1. Create Article Purchasing Data"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_2.6.1.1.1. Create Article Purchasing Data.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
'DataRowSet=7
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''----------------------Tcode-MM42 ----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextbox("Article","RMMW1-MATNR","",DT_MM42_0100_ARTICLE,False)
Call SetTextbox("Sales Org\.","RMMW1-VKORG","",DT_MM42_0100_SALES_ORG,False)
Call SetTextbox("Distr\. Channel","RMMW1-VTWEG","",DT_MM42_0100_DISTR_CHANNEL,False)
Call SetTextbox("Purchasing Org\.","RMMW1-EKORG","","",False)

Call TakeScreenShot()

Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100","Screen description","Purchasing",False)
Call TakeScreenShot()
Call PressEnter()
wait(2)
Call VerifyStatusBarMessageType("S") 
Call TakeScreenShot()
Call PressEnter()
wait(2)
Call TakeScreenShot()
'Call PressEnter()
wait(2)

Call SetTextbox("Rem\. Shelf Life","EINE-MHDRZ","",DT_MM42_2222_REM_SHELF_LIFE,False)
Call SetTextbox("Minimum Qty","EINE-MINBM","",DT_MM42_2222_MINIMUM_QTY,False)
Call SetTextbox("Pl\. Deliv\. Time","EINE-APLFZ","",DT_MM42_2222_PL_DELIV_TIME,False)
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Wait(2)
Call TakeScreenShot()

Call VerifyStatusBarMessageType("S")

Call VerifyStatusBar(DT_MM42_0100_CHECK_TEXT_OF_STATUSBAR)

Call LogOff()
Call FinalStatus ()
