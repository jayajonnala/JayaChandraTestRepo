'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_E10G07P02S01V01 Client Listing for BE10 
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_E10G07P02S01V01 Client for BE10"
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
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

'Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''''''--------TransactionCode-S_ALR_87012371----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextbox("Company Code","BR_BUKRS-LOW","",DT_S_ALR_87012371_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","BR_GJAHR-LOW","",DT_S_ALR_87012371_1000_FISCAL_YEAR,False)
Call SetTextbox("Posting date","BR_BUDAT-LOW","",ConvertDate(DT_S_ALR_87012371_1000_POSTING_DATE),False)
Call SetTextbox("to","BR_BUDAT-HIGH","",ConvertDate(DT_S_ALR_87012371_1000_TO),False)
Call SetTextbox("Ledger","BR_RLDNR-LOW","",DT_S_ALR_87012371_1000_LEDGER,False)
Call SetTextbox("Min./max.amount for tax report","GRENZBTR","","",False)
Call TakeScreenShot
Call ClickButton("%_BR_BELNR_%_APP_%-VALU_PUSH", False)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)
Call TakeScreenShot
Call PressEnter()
Call PressEnter()
Call SelectTab("TAB_STRIP", "Select Ranges", True)
Call SetTableData("SAPLALDBINTERVAL", "Lower limit", 1, "", "", DT_S_ALR_87012371_3020_TABLECELL_LOWER_LIMIT_0,True)
Call SetTableData("SAPLALDBINTERVAL", "Upper limit", 1, "", "", DT_S_ALR_87012371_3020_TABLECELL_UPPER_LIMIT_0,True)
Call TakeScreenShot
Call SelectTab("TAB_STRIP", "Select Single Values", True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "", "", DT_S_ALR_87012371_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "", "", DT_S_ALR_87012371_3010_TABLECELL_SINGLE_VALUE_1,True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 3, "", "", DT_S_ALR_87012371_3010_TABLECELL_SINGLE_VALUE_2,True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 4, "", "", "",True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 5, "", "", "",True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 6, "", "", "",True)
Call TakeScreenShot
Call ClickButton("Copy   \(F8\)", True)
Call ClickButton("%_S_MWSKZ_%_APP_%-VALU_PUSH", False)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC3)
Call TakeScreenShot
Call PressEnter()
Call PressEnter()
Call SetTableData("SAPLALDBSINGLE", "Single value",1, "", "",DT_S_ALR_87012371_3010_TABLECELL_SINGLE_VALUE_0_OCC1,True)
Call SetTableData("SAPLALDBSINGLE", "Single value",2, "", "",DT_S_ALR_87012371_3010_TABLECELL_SINGLE_VALUE_1_OCC1,True)
Call SetTableData("SAPLALDBSINGLE", "Single value",3, "", "",DT_S_ALR_87012371_3010_TABLECELL_SINGLE_VALUE_2_OCC1,True)
Call SetTableData("SAPLALDBSINGLE", "Single value",4, "", "",DT_S_ALR_87012371_3010_TABLECELL_SINGLE_VALUE_3,True)
Call SetTableData("SAPLALDBSINGLE", "Single value",5, "", "",DT_S_ALR_87012371_3010_TABLECELL_SINGLE_VALUE_4,True)
Call SetTableData("SAPLALDBSINGLE", "Single value",6, "", "",DT_S_ALR_87012371_3010_TABLECELL_SINGLE_VALUE_5,True)
Call SetTableData("SAPLALDBSINGLE", "Single value",7, "", "",DT_S_ALR_87012371_3010_TABLECELL_SINGLE_VALUE_6,True)
Call SetTableData("SAPLALDBSINGLE", "Single value",8, "", "",DT_S_ALR_87012371_3010_TABLECELL_SINGLE_VALUE_7,True)
Call SetTableData("SAPLALDBSINGLE", "Single value",9, "", "",DT_S_ALR_87012371_3010_SAPLALDBSINGLE_OCC2,True)
Call SetTableData("SAPLALDBSINGLE", "Single value",10, "", "",DT_S_ALR_87012371_3010_TABLECELL_SINGLE_VALUE_1_OCC2,True)
Call SetTableData("SAPLALDBSINGLE", "Single value",11, "", "",DT_S_ALR_87012371_3010_TABLECELL_SINGLE_VALUE_2_OCC1,True)
Call SetTableData("SAPLALDBSINGLE", "Single value",12, "", "",DT_S_ALR_87012371_3010_TABLECELL_SINGLE_VALUE_2_OCC2,True)
Call SetTableData("SAPLALDBSINGLE", "Single value",13, "", "",DT_S_ALR_87012371_3010_TABLECELL_SINGLE_VALUE_3_OCC1,True)
Call SetTableData("SAPLALDBSINGLE", "Single value",14, "", "",DT_S_ALR_87012371_3010_TABLECELL_SINGLE_VALUE_4_OCC1,True)
Call SetTableData("SAPLALDBSINGLE", "Single value",15, "", "",DT_S_ALR_87012371_3010_TABLECELL_SINGLE_VALUE_5_OCC1,True)
Call SetTableData("SAPLALDBSINGLE", "Single value",16, "", "",DT_S_ALR_87012371_3010_TABLECELL_SINGLE_VALUE_6_OCC1,True)
Call SetTableData("SAPLALDBSINGLE", "Single value",17, "", "",DT_S_ALR_87012371_3010_TABLECELL_SINGLE_VALUE_7_OCC1,True)
Call SetTableData("SAPLALDBSINGLE", "Single value",18, "", "",DT_S_ALR_87012371_3010_SAPLALDBSINGLE_OCC3,True)
Call SetTableData("SAPLALDBSINGLE", "Single value",19, "", "",DT_S_ALR_87012371_3010_SAPLALDBSINGLE_OCC4,True)
Call SetTableData("SAPLALDBSINGLE", "Single value",20, "", "",DT_S_ALR_87012371_3010_TABLECELL_SINGLE_VALUE_1_OCC3,True)
Call SetTableData("SAPLALDBSINGLE", "Single value",21, "", "",DT_S_ALR_87012371_3010_TABLECELL_SINGLE_VALUE_2_OCC3,True)
Call SetTableData("SAPLALDBSINGLE", "Single value",22, "", "",DT_S_ALR_87012371_3010_TABLECELL_SINGLE_VALUE_3_OCC2,True)
Call SetTableData("SAPLALDBSINGLE", "Single value",23, "", "",DT_S_ALR_87012371_3010_TABLECELL_SINGLE_VALUE_4_OCC2,True)
Call SetTableData("SAPLALDBSINGLE", "Single value",24, "", "",DT_S_ALR_87012371_3010_TABLECELL_SINGLE_VALUE_5_OCC2,True)
Call SetTableData("SAPLALDBSINGLE", "Single value",25, "", "",DT_S_ALR_87012371_3010_TABLECELL_SINGLE_VALUE_6_OCC2,True)
Call SetTableData("SAPLALDBSINGLE", "Single value",26, "", "",DT_S_ALR_87012371_3010_TABLECELL_SINGLE_VALUE_7_OCC2,True)
Call SetTableData("SAPLALDBSINGLE", "Single value",27, "", "",DT_S_ALR_87012371_3010_TABLECELL_SINGLE_VALUE_7_OCC3,True)
Call TakeScreenShot
Call ClickButton("Copy   \(F8\)", True)
Call ClickButton("Execute   \(F8\)", False)
Call TakeScreenShot
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC6)
Call PressEnter()
Call PressEnter()
Call ClickButton("Back   \(F3\)", False)
Call ClickButton("Next page   \(Page down\)", False)
Call ClickButton("Next page   \(Page down\)", False)
Call ClickButton("Next page   \(Page down\)", False)
Call ClickButton("Next page   \(Page down\)", False)
Call ClickButton("Next page   \(Page down\)", False)
Call ClickButton("Next page   \(Page down\)", False)
Call TakeScreenShot
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012371_0120_CHECK_TEXT_OF_NO_NAME, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012371_0120_CHECK_TEXT_OF_NO_NAME_OCC1, 1)
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
