
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_159_Update Account Completion Table (All Opcos)_TASE
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)


gstrTestCaseName = "Test_06-01-02-02-02-Check MSI Net Orders in SAP_SO"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
Call ClickButtonIfExist("Continue (Enter)",True)

'-----------------------------------------ME23N---------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Purchase Order No\.","RV45S-BSTNK","",DT_VA03_0102_PURCHASE_ORDER_NO,False)
Call ClickButton("Search",False)
Call TakeScreenShot
'Call ClickLabel("20\.06\.2022","",True)
Call ClickButton("Copy   \(Enter\)",True)
Call ClickButton("Display header details",False)
Call TakeScreenShot
Call VerifyTextboxContent("Pricing Date","VBKD-PRSDT","",DT_VA03_4301_CHECK_TEXT_OF_PRICING_DATE,False)
Call VerifyTextboxContent("Order Type","VBAK-AUART","",DT_VA03_4301_CHECK_TEXT_OF_ORDER_TYPE,False)
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call VerifyTableCellContent(1,"Article","SAPMV45ATCTRL_U_ERF_AUFTRAG",DT_VA03_4900_CHECK_TEXT_OF_TABLECELL_ARTICLE_0)
Call VerifyTableCellContent(2,"Article","SAPMV45ATCTRL_U_ERF_AUFTRAG",DT_VA03_4900_CHECK_TEXT_OF_TABLECELL_ARTICLE_1)
Call TakeScreenShot
Call VerifyTableCellContent(1,"Order Quantity","SAPMV45ATCTRL_U_ERF_AUFTRAG",DT_VA03_4900_CHECK_TEXT_OF_TABLECELL_ORDER_QUANTITY_0)
Call VerifyTableCellContent(2,"Order Quantity","SAPMV45ATCTRL_U_ERF_AUFTRAG",DT_VA03_4900_CHECK_TEXT_OF_TABLECELL_ORDER_QUANTITY_1)
'Call VerifyTableCellContent(1,"Site","SAPMV45ATCTRL_U_ERF_AUFTRAG",DT_VA03_4900_CHECK_TEXT_OF_TABLECELL_SITE_0)
Call VerifyTableCellContent(2,"Site","SAPMV45ATCTRL_U_ERF_AUFTRAG",DT_VA03_4900_CHECK_TEXT_OF_TABLECELL_SITE_1)
Call GetTextboxValue("VBAK-VBELN","","DT_VA03_4021_CHECK_TEXT_OF_SALESORDERRETAIL_AUT_Output",False)
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet("DT_VA03_4021_CHECK_TEXT_OF_SALESORDERRETAIL_AUT_Output",DT_VA03_4021_CHECK_TEXT_OF_SALESORDERRETAIL_AUT_Output)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call LogOff()
Call FinalStatus()

