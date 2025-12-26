'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_02-06-05-01-20-Manual Wholesale Price calculation_multiple_EA_TASE
'.................Test Scenario: AT_04-06-02-06-04-Vendor returns-DC Drinks-No PO-Tied Empties
'.................TCode: ME21N,ME23N
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


gstrTestCaseName = "TC1_Test_02-06-05-01-20-Manu EA"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.

'-------------------------------------------------------------------------------------------
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)'.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'---------------------VK11----------------------
Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter()
Call TakeScreenShot()

Call SetTextBox("Condition Type","RV13A-KSCHL","",DT_VK11_0100_CONDITION_TYPE,False)
Call TakeScreenShot()
''''Call PressEnter()

Call ClickButton("Key Combination   \(Shift\+F5\)",False)
Call SelectRadioButton("RV130-SELKZ","Sales Org\./Dist\. Channel/Price List/Article/Sales Unit",True)
Call TakeScreenShot()
Call ClickButton("Choose   \(Enter\)",True)

Call SetTextBox("Sales Organization","KOMG-VKORG","",DT_VK11_1155_SALES_ORGANIZATION,False)
Call SetTextBox("Distribution Channel","KOMG-VTWEG","",DT_VK11_1155_DISTRIBUTION_CHANNEL,False)
Call SetTextBox("Price List","KOMG-PLTYP","",DT_VK11_1155_PRICE_LIST,False)
Call TakeScreenShot()
Call PressEnter()

Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Article","1","","",DT_VK11_1155_TABLECELL_ARTICLE_0,False)
Call PressEnter()
Call TakeScreenShot()
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Sales Unit","1","","",DT_VK11_1155_TABLECELL_SALES_UNIT_0,False)
Call PressEnter()
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Amount","1","","",DT_VK11_1155_TABLECELL_AMOUNT_0,False)
Call PressEnter()
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call TakeScreenShot()
Call VerifyStatusBar(DT_VK11_1155_CHECK_TEXT_OF_STATUSBAR)

'--------------------------------------VKP5----------------------'
Call SetTcode(DT_VK11_1155_OKCD)
Call PressEnter()
Call TakeScreenShot()

Call SetTextBox("Article","S_MATNR-LOW","",DT_VK11_1000_ARTICLE,False)
Call SetTextBox("Sales Organization","S_VKORG-LOW","",DT_VK11_1000_SALES_ORGANIZATION,False)
Call SetTextBox("Distribution Channel","S_VTWEG-LOW","",DT_VK11_1000_DISTRIBUTION_CHANNEL,False)
Call PressEnter()
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)

Call SelectCheckboxNoLabel(0,"ON",False)
Call TakeScreenShot()
Call ClickButton("Sales   \(F6\)",False)
Call TakeScreenShot()

Call VerifyTextBoxContent("Material","KOMP-MATNR","",DT_VK11_9000_CHECK_TEXT_OF_ARTICLE,False)
Call VerifyTableCellContent(1,"CnTy","SAPLV69ATCTRL_KONDITIONEN",DT_VK11_6201_CHECK_TEXT_OF_TABLECELL_CNTY_0)

Call GetTableCellData("SAPLV69ATCTRL_KONDITIONEN","NumCCo",6,"","","DT_NUM",False)
Call GetTableCellData("SAPLV69ATCTRL_KONDITIONEN","Amount",15,"","","DT_TAX",False)

Call VerifyTableCellContent(1,"Amount","SAPLV69ATCTRL_KONDITIONEN",DT_VK11_6201_CHECK_TEXT_OF_TABLECELL_AMOUNT_0)
Call VerifyTableCellContent(1,"Crcy","SAPLV69ATCTRL_KONDITIONEN",DT_VK11_6201_CHECK_TEXT_OF_TABLECELL_CRCY_0)


' VerifyTableCellContent(tableRowNumber, tableColumnName, tableName, expectedValue)
VerifyTableCellContent
Call TakeScreenShot()
Call LogOff()
Call FinalStatus()

