

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_2.2.1.3.2. Block Vendor Account Group ZVXT - External Trade Vend _TASE
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

gstrTestCaseName = "Test_2.2.1.3.2. Block Vendor Account Group ZVXT - External Trade Vend _TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\P3\DS\RETAIL\DT_2.2.1.3.2. Block Vendor Account Group ZVXT - External Trade Vend_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''
''----------------------Tcode ZMDPU_INFOREC_COPY----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'Call SetTextbox("Vendor","RF02K-LIFNR","",DT_XK05_0500_VENDOR,False)
Call SetTextboxNoLabel("RF02K-LIFNR","",DT_XK05_0500_VENDOR,False)
Call SetTextbox("Company Code","RF02K-BUKRS","",DT_XK05_0500_COMPANY_CODE,False)
Call SetTextbox("Purch\. Organization","RF02K-EKORG","",DT_XK05_0500_PURCH_ORGANIZATION,False)
Call TakeScreenShot
Call PressEnter() 

Call SetTextbox("Block function","LFA1-SPERQ","",DT_XK05_0510_BLOCK_FUNCTION,False)
Call SelectCheckbox("LFB1-SPERR","1","ON",False)
Call SelectCheckbox("LFM1-SPERM","1","ON",False)
Call TakeScreenShot

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call VerifyStatusBar(DT_XK05_0500_CHECK_TEXT_OF_STATUSBAR)
Call PressEnter() 

Call SelectCheckbox("LFB1-SPERR","1","OFF",False)
Call SelectCheckbox("LFM1-SPERM","1","OFF",False)
Call SetTextbox("Block function","LFA1-SPERQ","","",False)
Call TakeScreenShot

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call VerifyStatusBar(DT_XK05_0500_CHECK_TEXT_OF_STATUSBAR)


Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
