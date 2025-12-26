
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRI_02_053-Eshop comission manual in SAP - correct values
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
'.................Test Script Name :Test_S2A_PRI_02_053-Eshop comission manual in SAP - correct values
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_S2A_PRI_02_053-Eshop  - correct values"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_S2A_PRI_02_053-Eshop comission manual in SAP - correct values.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)

Call Login(DT_SAPUSER,DT_SAPPASSWORD)

Call PressEnter()

Call SetTcode(DT_SAPTRANSACTIONCODE)     

Call PressEnter()   
Call SetTextbox("Condition Type","RV13A-KSCHL","",DT_VK11_0100_CONDITION_TYPE,False)
Call PressEnter()
Call SelectRadioButton("RV130-SELKZ", "Sales Org./Dist. Channel/Price List/Article/Sales Unit", True)
Call ClickButtonIfExist("Choose   \(Enter\)",True)
Call SetTextbox("Sales Organization","KOMG-VKORG","",DT_VK11_1155_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution Channel","KOMG-VTWEG","",DT_VK11_1155_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Price List","KOMG-PLTYP","",DT_VK11_1155_PRICE_LIST,False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Article", 1, "<NA>", "<NA>", DT_VK11_1155_TABLECELL_ARTICLE_0, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Article", 2, "<NA>", "<NA>", DT_VK11_1155_TABLECELL_ARTICLE_1, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Article", 3, "<NA>", "<NA>", DT_VK11_1155_TABLECELL_ARTICLE_2, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Article", 4, "<NA>", "<NA>", DT_VK11_1155_TABLECELL_ARTICLE_3, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Article", 5, "<NA>", "<NA>", DT_VK11_1155_TABLECELL_ARTICLE_4, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Article", 6, "<NA>", "<NA>", DT_VK11_1155_TABLECELL_ARTICLE_5, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Article", 7, "<NA>", "<NA>", DT_VK11_1155_TABLECELL_ARTICLE_6, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Sales unit", 1, "<NA>", "<NA>", DT_VK11_1155_TABLECELL_SALES_UNIT_0, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Sales unit", 2, "<NA>", "<NA>", DT_VK11_1155_TABLECELL_SALES_UNIT_1, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Sales unit", 3, "<NA>", "<NA>", DT_VK11_1155_TABLECELL_SALES_UNIT_2, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Sales unit", 4, "<NA>", "<NA>", DT_VK11_1155_TABLECELL_SALES_UNIT_3, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Sales unit", 5, "<NA>", "<NA>", DT_VK11_1155_TABLECELL_SALES_UNIT_4, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Sales unit", 6, "<NA>", "<NA>", DT_VK11_1155_TABLECELL_SALES_UNIT_5, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Sales unit", 7, "<NA>", "<NA>", DT_VK11_1155_TABLECELL_SALES_UNIT_6, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Amount", 1, "<NA>", "<NA>", DT_VK11_1155_TABLECELL_AMOUNT_0, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Amount", 2, "<NA>", "<NA>", DT_VK11_1155_TABLECELL_AMOUNT_1, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Amount", 3, "<NA>", "<NA>", DT_VK11_1155_TABLECELL_AMOUNT_2, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Amount", 4, "<NA>", "<NA>", DT_VK11_1155_TABLECELL_AMOUNT_3, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Amount", 5, "<NA>", "<NA>", DT_VK11_1155_TABLECELL_AMOUNT_4, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Amount", 6, "<NA>", "<NA>", DT_VK11_1155_TABLECELL_AMOUNT_5, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Amount", 7, "<NA>", "<NA>", DT_VK11_1155_TABLECELL_AMOUNT_6, False)
Call PressEnter()     
Call TakeScreenShot
Call SelectRowGuiTable("SAPMV13ATCTRL_FAST_ENTRY", "Article", DT_VK11_1155_TABLECELL_ARTICLE_6, False)
Call ClickButton("Change validity",False)
Call SetTextbox("Valid to","RV13A-DATBI","",ConvertDate(DT_VK11_0210_VALID_TO),True)
Call TakeScreenShot
Call ClickButton("Copy Date   \(Enter\)",True)
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)

'''''--------TransactionCode-/nvk13----------''''

Call SetTcode(DT_VK11_1155_OKCD)     
Call PressEnter()     

Call SetTextbox("Condition Type","RV13A-KSCHL","",DT_VK11_0100_CONDITION_TYPE_OCC1,False)
Call ClickButton("Condition Information   \(Shift\+F4\)",False) 
Call SetTextbox("Article","F001-LOW","",DT_VK11_1000_ARTICLE,False)
Call SetTextbox("Sales Organization","F002-LOW","",DT_VK11_1000_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution Channel","F003-LOW","",DT_VK11_1000_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Price List","F007-LOW","",DT_VK11_1000_PRICE_LIST,False)
Call SetTextbox("from / on","SEL_DATE","",ConvertDate(DT_VK11_1000_FROM__ON),False)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call SelectCheckboxNoLabel(0, DT_VK11_0120_NO_NAME, False)
Call ClickButton("Display   \(F5\)",False)
Call TakeScreenShot
Call VerifyTableCellContent(1, "Article", "SAPMV13ATCTRL_FAST_ENTRY", DT_VK11_1073_CHECK_TEXT_OF_TABLECELL_ARTICLE_0)
Call VerifyTableCellContent(1, "Amount", "SAPMV13ATCTRL_FAST_ENTRY", DT_VK11_1073_CHECK_TEXT_OF_TABLECELL_AMOUNT_0)
Call VerifyTableCellContent(1, "Sales unit", "SAPMV13ATCTRL_FAST_ENTRY", DT_VK11_1073_CHECK_TEXT_OF_TABLECELL_SALES_UNIT_0)
Call ClickButton("Back   \(F3\)",False)
Call SelectCheckboxNoLabel(1, DT_VK11_0120_NO_NAME_OCC1, False)
Call ClickButton("Display   \(F5\)",False)
Call TakeScreenShot
Call VerifyTableCellContent(1, "Article", "SAPMV13ATCTRL_FAST_ENTRY", DT_VK11_1155_CHECK_TEXT_OF_TABLECELL_ARTICLE_0)
Call VerifyTableCellContent(1, "Amount", "SAPMV13ATCTRL_FAST_ENTRY", DT_VK11_1155_CHECK_TEXT_OF_TABLECELL_AMOUNT_0)
Call VerifyTableCellContent(1, "Sales unit", "SAPMV13ATCTRL_FAST_ENTRY", DT_VK11_1155_CHECK_TEXT_OF_TABLECELL_SALES_UNIT_0)
Call VerifyTableCellContent(1, "Valid From", "SAPMV13ATCTRL_FAST_ENTRY", ConvertDate(DT_VK11_1155_CHECK_TEXT_OF_TABLECELL_VALID_FROM_0))
Call VerifyTableCellContent(1, "Valid to", "SAPMV13ATCTRL_FAST_ENTRY", ConvertDate(DT_VK11_1155_CHECK_TEXT_OF_TABLECELL_VALID_TO_0))


Call LogOff()

Call FinalStatus()





'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [8,3437477]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//




