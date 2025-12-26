
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Distribution process to DS stores - SW35 (Consumables and1
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
'.................Test Script Name :Test_Distribution process to DS stores - SW35 (Consumables and1
'.................Author : TCS 	   
'................ Creation Date    :16th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Distribution process to DS stores - SW35 (Consumables and1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\P3\DS\RETAIL\DT_Distribution process to DS stores - SW35 (Consumables and1_TASE.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode VL02N----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
Call TakeScreenShot()

Call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_VL02N_4004_OUTBOUND_DELIVERY,FALSE)
Call TakeScreenShot()
Call PressEnter()
Wait(3)
Call SetTextbox("Document Date","LIKP-BLDAT","",Replace(DT_VL02N_1502_DOCUMENT_DATE,"/","."),False)
Call TakeScreenShot()

Call ClickButtonIfExist("Post Goods Issue   \(Shift\+F8\)",False)
wait(2)
Call TakeScreenShot()
Call VerifyStatusBarMessageType(DT_EXPECTED_MESSAGE)

Call VerifyStatusBar(DT_VL02N_4004_CHECK_TEXT_OF_STATUSBAR)
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

