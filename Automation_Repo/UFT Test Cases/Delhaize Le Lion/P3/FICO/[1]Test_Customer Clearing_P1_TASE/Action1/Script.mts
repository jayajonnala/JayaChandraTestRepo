'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_Customer Clearing_PRE1
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Customer Clearing_P1"
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
Call EndDateof445PeriodByDate(DT_TODAY,"DT_ENDING_DATE_OF_CURRENT_PERIOD")

''''''--------------login----------------'''''

Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''--------TransactionCode-ZFIAR_AFF_CLEARING----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE) 

Call SetTextbox("Document Date","P_BLDAT","",ConvertDate(DT_ZFIAR_AFF_CLEARING_1000_DOCUMENT_DATE),False)
Call SetTextbox("Posting Date","P_BUDAT","",ConvertDate(DT_ZFIAR_AFF_CLEARING_1000_POSTING_DATE),False)
Call SetTextbox("Company Code","S_BUKRS-LOW","",DT_ZFIAR_AFF_CLEARING_1000_COMPANY_CODE,False)
Call SetTextbox("Document Header Text","P_BKTXT","",DT_ZFIAR_AFF_CLEARING_1000_DOCUMENT_HEADER_TEXT,False)
Call SetTextbox("Open Items at Key Date","P_KEYDT","",ConvertDate(DT_ZFIAR_AFF_CLEARING_1000_OPEN_ITEMS_AT_KEY_DATE),False)
Call SetTextbox("Branch account","S_FILKD-LOW","",DT_ZFIAR_AFF_CLEARING_1000_BRANCH_ACCOUNT,False)
Call SetTextbox("Customer","S_KUNNR-LOW","",DT_ZFIAR_AFF_CLEARING_1000_CUSTOMER,False)
Call SetTextbox("Fiscal Year","S_GJAHR-LOW","",DT_ZFIAR_AFF_CLEARING_1000_FISCAL_YEAR,False)
Call ClickButtonIfExist("%_S_BLART_%_APP_%-VALU_PUSH",False)
Call SelectTab("TAB_STRIP", "Exclude Single Values", True)
Call SetTableData("SAPLALDBSINGLE_E", "Single value", 1, "", "",DT_ZFIAR_AFF_CLEARING_3030_TABLECELL_SINGLE_VALUE_0, True)
Call SetTableData("SAPLALDBSINGLE_E", "Single value", 2, "", "",DT_ZFIAR_AFF_CLEARING_3030_TABLECELL_SINGLE_VALUE_1 , True)
Call SetTableData("SAPLALDBSINGLE_E", "Single value", 3, "", "", DT_ZFIAR_AFF_CLEARING_3030_TABLECELL_SINGLE_VALUE_2, True)
Call SetTableData("SAPLALDBSINGLE_E", "Single value", 4, "", "",DT_ZFIAR_AFF_CLEARING_3030_TABLECELL_SINGLE_VALUE_3 , True)
Call SetTableData("SAPLALDBSINGLE_E", "Single value", 5, "", "",DT_ZFIAR_AFF_CLEARING_3030_TABLECELL_SINGLE_VALUE_4 , True)
Call SetTableData("SAPLALDBSINGLE_E", "Single value", 6, "", "",DT_ZFIAR_AFF_CLEARING_3030_TABLECELL_SINGLE_VALUE_5 , True)
Call TakeScreenShot()
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot()
Call VerifyTextBoxContent("Information Message", "MESSTXT1", "",lcase(DT_ZFIAR_AFF_CLEARING_0010_CHECK_TEXT_OF_MESSTXT1), True)
Call GetTextboxValue("MESSTXT1", "","DT_OP_FB01_0010_CHECK_TEXT_OF_MESSTXT1", True) 
GetRowNo =2
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
