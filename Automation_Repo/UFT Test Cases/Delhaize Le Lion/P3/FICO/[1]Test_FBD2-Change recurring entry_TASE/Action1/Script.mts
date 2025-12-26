'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_FBD2-Change recurring entry
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_FBD2-Change recurring entry"
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

''''''--------TransactionCode-FBD2----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Document Number","RF05L-BELNR","",DT_FBD2_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FBD2_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FBD2_0100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter() 
Call VerifyGridCellContent("", 1, "Amount", "", DT_FBD2_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("", 2, "Amount", "", DT_FBD2_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call DoubleClickGuiGridCell("", "", 1, "Account", False)
Call TakeScreenShot
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FBD2_0300_AMOUNT,False)
Call TakeScreenShot
Call PressEnter()
Call VerifyTextBoxContent("Amount","BSEG-WRBTR", "", DT_FBD2_0300_CHECK_TEXT_OF_AMOUNT, False)
Call ClickButton("Display Next Item   \(Shift\+F7\)",False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FBD2_0300_AMOUNT_OCC1,False)
Call PressEnter()
Call VerifyTextBoxContent("Amount","BSEG-WRBTR", "",DT_FBD2_0300_CHECK_TEXT_OF_AMOUNT_OCC1, False)
Call PressEnter()
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call VerifyStatusBar(DT_FBD2_0100_CHECK_TEXT_OF_STATUSBAR)
Call SetTextbox("Document Number","RF05L-BELNR","",DT_FBD2_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FBD2_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FBD2_0100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter() 
Call VerifyGridCellContent("", 1, "Amount", "", DT_FBD2_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET_OCC1)
Call VerifyGridCellContent("", 2, "Amount", "", DT_FBD2_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET_OCC1)
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
