'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_Customer Clearing_P4  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Customer Clearing_P4"
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

''''''--------TransactionCode-/FB03----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB03_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB03_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FB03_0100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()
Call VerifyGridCellContent("", 1, "Posting Key","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "Posting Key","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 1, "Amount","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("", 2, "Amount","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call TakeScreenShot
Call ClickButtonIfExist("Document Display: General Ledger View   \(Ctrl\+F9\)",False)
Call VerifyGridCellContent("", 1, "Posting Key","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OCC1)
Call VerifyGridCellContent("", 2, "Posting Key","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL_OCC1)
Call VerifyGridCellContent("", 3, "Posting Key","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL_OCC1)
Call VerifyGridCellContent("", 4, "Posting Key","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BSCHL_OCC1)
Call VerifyGridCellContent("", 1, "Amount","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET_OCC1)
Call VerifyGridCellContent("", 2, "Amount","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET_OCC1)
Call VerifyGridCellContent("", 1, "Amount","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AZBET_OCC1)
Call VerifyGridCellContent("", 2, "Amount","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_AZBET_OCC1)
Call TakeScreenShot
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
