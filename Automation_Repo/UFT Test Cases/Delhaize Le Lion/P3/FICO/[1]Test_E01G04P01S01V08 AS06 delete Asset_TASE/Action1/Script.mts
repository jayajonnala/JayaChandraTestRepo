'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_02-04-01-05-03-Create new assortment-BASA  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_E01G04P01S01V08 AS06 del As"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\DLL_P3\Data\TASE_DT_02-04-01-05-03-Create new assortment-BASA.xls"

If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	datatable_row= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''''''--------TransactionCode-AS06----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS06_0100_COMPANY_CODE,False)
Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS06_0100_ASSET,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AS06_0100_SUBNUMBER,False)
Call TakeScreenShot
Call PressEnter()
Call SelectRadioButton("RA02S-XPHYS","Physically delete asset", False)
Call TakeScreenShot
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call VerifyTextBoxNoLabelContent("SPOP-TEXTLINE1","", lcase(DT_AS06_0100_CHECK_TEXT_OF_SPOPTEXTLINE1),True)
Call ClickButtonIfExist("Yes",True)
Call TakeScreenShot
Call VerifyStatusBar(DT_AS06_0100_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot

''''''--------TransactionCode-AW01----------''''

Call SetTcode(DT_EXPECTEDTRANSACTIONCODE_OCC1)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)

Call VerifyStatusBar(DT_AS06_0100_CHECK_TEXT_OF_STATUSBAR_OCC1)
Call TakeScreenShot

''''''--------TransactionCode-AS03----------''''
Call SetTcode(DT_EXPECTEDTRANSACTIONCODE_OCC2)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS06_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS06_0100_ASSET_OCC1,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AS06_0100_SUBNUMBER_OCC1,False)
Call TakeScreenShot
Call PressEnter()
Call VerifyStatusBar(DT_AS06_0100_CHECK_TEXT_OF_STATUSBAR_OCC2)
Call TakeScreenShot
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
