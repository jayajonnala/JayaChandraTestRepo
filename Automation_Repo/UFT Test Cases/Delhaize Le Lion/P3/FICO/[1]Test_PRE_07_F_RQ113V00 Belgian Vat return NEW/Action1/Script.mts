'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_PRE_07_F_RQ113V00 Belgian Vat return NEW  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_PRE_07_F_RQ113V00 Belgian Vat return NEW"
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

'''''--------TransactionCode-S_ALR_87012357----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call ClickButton("Get Variant...   \(Shift\+F5\)", False)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call SetTextbox("Variant","V-LOW","",DT_S_ALR_87012357_0600_GRIDCELL_42_VARIANT_NAME,True)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)", True)
wait 5

Call SetTextbox("Company Code","BR_BUKRS-LOW","",DT_S_ALR_87012357_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","BR_GJAHR-LOW","","2024",False)
Call SetTextbox("Posting date","BR_BUDAT-LOW","",ConvertDate(DT_S_ALR_87012357_1000_POSTING_DATE),False)
Call SetTextbox("to","BR_BUDAT-HIGH","",ConvertDate(DT_S_ALR_87012357_1000_TO),False)
Call TakeScreenShot
'Call ClickButton("PUSHB_O2", False)
Call SetTextbox("Due Date of the Tax Payable","PAR_FDAT","",ConvertDate(DT_S_ALR_87012357_1000_DUE_DATE_OF_THE_TAX_PAYABLE),False)
Call SetTextbox("Posting date","PAR_BDAT","",ConvertDate(DT_S_ALR_87012357_1000_POSTING_DATE),False)
Call TakeScreenShot
Call ClickButton("%_BR_BELNR_%_APP_%-VALU_PUSH", False)
Call SelectTab("TAB_STRIP", "Select Ranges", True)
Call SetTableData("SAPLALDBINTERVAL", "Lower limit", 1, "", "", DT_S_ALR_87012357_3020_TABLECELL_LOWER_LIMIT_0,True)
Call SetTableData("SAPLALDBINTERVAL", "Upper limit", 1, "", "", DT_S_ALR_87012357_3020_TABLECELL_UPPER_LIMIT_0,True)
Call SetTableData("SAPLALDBINTERVAL", "Lower limit", 2, "", "", DT_S_ALR_87012357_3020_TABLECELL_LOWER_LIMIT_1,True)
Call SetTableData("SAPLALDBINTERVAL", "Upper limit", 2, "", "", DT_S_ALR_87012357_3020_TABLECELL_UPPER_LIMIT_1,True)
Call SetTableData("SAPLALDBINTERVAL", "Lower limit", 3, "", "", DT_S_ALR_87012357_3020_TABLECELL_LOWER_LIMIT_2,True)
Call SetTableData("SAPLALDBINTERVAL", "Upper limit", 3, "", "", DT_S_ALR_87012357_3020_TABLECELL_UPPER_LIMIT_2,True)
Call TakeScreenShot
Call SelectTab("TAB_STRIP", "Select Single Values", True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "", "", DT_S_ALR_87012357_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "", "", DT_S_ALR_87012357_3010_TABLECELL_SINGLE_VALUE_1,True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 3, "", "", DT_S_ALR_87012357_3010_TABLECELL_SINGLE_VALUE_2,True)
Call TakeScreenShot
Call ClickButton("Copy   \(F8\)", True)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)", False)
Call ClickButton("Last Page   \(Ctrl\+Page down\)", False)
Call ClickButton("Previous Page   \(Page up\)", False)
Call TakeScreenShot
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_NO_NAME, 1)
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
