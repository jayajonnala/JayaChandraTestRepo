
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_01_01_015-Create Single Article SAP (ZHAW)
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

gstrTestCaseName = "Test_MD_01_07_119-Manage Retail  - VKP0 - JDA - ZHAW"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\FICO\TASE_DT_AA001 - Scrapping Asset partial Deduct with NBV, IFRS equal Loca.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'----------------------Tcode vk11 ----------------------------
'Enter the transaction code
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     
'Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextBox("Condition Type","RV13A-KSCHL","",DT_VK11_0100_CONDITION_TYPE,False)
''Call SetComboByKey("RV13A-KSCHL",DT_MM41_0100_ARTICLE_TYPE)
Call PressEnter()
' SelectRadioButton(radiobuttonName, radiobuttonAttachedtext, blnIsItPopup)
Call SelectRadioButton("RV130-SELKZ","Article per SOrg/DstCh",True)
Call ClickButton("Choose   \(Enter\)",True)
Call SetTextBox("Sales Organization","KOMG-VKORG","",DT_VK11_1073_SALES_ORGANIZATION,False)
Call SetTextBox("Distribution Channel","KOMG-VTWEG","",DT_VK11_1073_DISTRIBUTION_CHANNEL,False)
' SetTableData(tableName, columnName, rowNumber, refColumnName, refCellValue, cellValue, blnIsItPopup)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Article",1,"","",DT_VK11_1073_TABLECELL_ARTICLE_0,False)
Call PressEnter()
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Sales Unit",1,"","",DT_VK11_1073_TABLECELL_SALES_UNIT_0,False)
Call PressEnter()
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Amount",1,"","",DT_VK11_1073_TABLECELL_AMOUNT_0,False)
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Enter   \(F5\)",True)
''Call GetStatusBar("item2","DT_VK11_1073_CHECK_TEXT_OF_STATUSBAR")
Call VerifyStatusBar(DT_VK11_1073_CHECK_TEXT_OF_STATUSBAR)

'''''----------------------Tcode vk13 ----------------------------
''''Enter the transaction code
Call SetTcode(DT_VK11_1073_OKCD) 
Call PressEnter() 
 
Call SetTextBox("Condition Type","RV13A-KSCHL","",DT_VK11_0100_CONDITION_TYPE_OCC1,False)
'Call SetComboByKey("RV13A-KSCHL",DT_VK11_0100_CONDITION_TYPE_OCC1)
Call PressEnter()
' SelectRadioButton(radiobuttonName, radiobuttonAttachedtext, blnIsItPopup)
Call SelectRadioButton("RV130-SELKZ","Article per SOrg/DstCh",True)
Call ClickButton("Choose   \(Enter\)",True)
Call SetTextBox("Sales Organization","F001","",DT_VK11_1000_SALES_ORGANIZATION,False)
Call SetTextBox("Distribution Channel","F002","",DT_VK11_1000_DISTRIBUTION_CHANNEL,False)
Call SetTextBox("Article","F003-LOW","",DT_VK11_1000_ARTICLE,False)
Call SetTextBox("Valid on","SEL_DATE","",ConvertDate(DT_VK11_1000_VALID_ON),False)
Call ClickButton("Execute   \(F8\)",False)
' VerifyTableCellContent(tableRowNumber, tableColumnName, tableName, expectedValue)
Call VerifyTableCellContent(1,"Article","SAPMV13ATCTRL_FAST_ENTRY",DT_VK11_1073_CHECK_TEXT_OF_TABLECELL_ARTICLE_0)
Call VerifyTableCellContent(1,"Sales Unit","SAPMV13ATCTRL_FAST_ENTRY",DT_VK11_1073_CHECK_TEXT_OF_TABLECELL_SALES_UNIT_0)

Call VerifyTableCellContent(1,"Amount","SAPMV13ATCTRL_FAST_ENTRY",Cstr(DT_VK11_1073_CHECK_TEXT_OF_TABLECELL_AMOUNT_0&" "))
Call VerifyTableCellContent(1,"Valid From","SAPMV13ATCTRL_FAST_ENTRY",ConvertDate(DT_VK11_1073_CHECK_TEXT_OF_TABLECELL_VALID_FROM_0))
' VerifyTextBoxContent(textboxAttachedText, textboxName, textboxIndex, expectedValue, blnIsItPopup)
Call VerifyTextBoxContent("Sales Organization","KOMG-VKORG","",DT_VK11_1073_CHECK_TEXT_OF_SALES_ORGANIZATION,False)
' VerifyTextBoxContent(textboxAttachedText, textboxName, textboxIndex, expectedValue, blnIsItPopup)
Call VerifyTextBoxContent("Distribution Channel","KOMG-VTWEG","",DT_VK11_1073_CHECK_TEXT_OF_DISTRIBUTION_CHANNEL,False)





'Call SetTextbox("Mdse Catgry","RMMW1-MATKL","",DT_MM41_0100_MDSE_CATGRY,False)
'Call SetComboByKey("RMMW1-ATTYP",DT_MM41_0100_ARTL_CATEGORY)
'Call SetTextbox("Purchasing Org\.","RMMW1-EKORG","",DT_MM41_0100_PURCHASING_ORG,False)
'Call SetTextbox("Vendor","RMMW1-LIFNR","",DT_MM41_0100_VENDOR,False)
'
'Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100","Screen description","Basic Data",False)
'Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100","Screen description","Purchasing",False)
'Call TakeScreenShot()
'Call PressEnter() 
'
'Call SetTableData("SAPLMGD2TC_ME_8022", "AUoM", 1, "", "", DT_MM41_8022_TABLECELL_AUOM_0, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "AUoM", 2, "", "", DT_MM41_8022_TABLECELL_AUOM_1, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "AUoM", 3, "", "", DT_MM41_8022_TABLECELL_AUOM_2, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "AUoM", 4, "", "", DT_MM41_8022_TABLECELL_AUOM_3, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "Number", 2, "", "", DT_MM41_8022_TABLECELL_NUMBER_1, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "Number", 3, "", "", DT_MM41_8022_TABLECELL_NUMBER_2, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "Number", 4, "", "", DT_MM41_8022_TABLECELL_NUMBER_3, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "OUn", 2, "", "", "ON", False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "D/I", 2, "", "", "ON", False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "SUn", 1, "", "", "ON", False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "LUn", 2, "", "",DT_MM41_8022_TABLECELL_LUN_1, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "LUn", 3, "", "",DT_MM41_8022_TABLECELL_LUN_2, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "LUn", 4, "", "",DT_MM41_8022_TABLECELL_LUN_3, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "GTIN", 1, "", "",DT_MM41_8022_TABLECELL_GTIN_0, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "Gross Weight", 1, "", "",DT_MM41_8022_TABLECELL_GROSS_WEIGHT_0, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "Gross Weight", 2, "", "",DT_MM41_8022_TABLECELL_GROSS_WEIGHT_1, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "Gross Weight", 3, "", "",DT_MM41_8022_TABLECELL_GROSS_WEIGHT_2, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "Gross Weight", 4 ,"", "",DT_MM41_8022_TABLECELL_GROSS_WEIGHT_3, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "Net Weight", 1, "", "",DT_MM41_8022_TABLECELL_NET_WEIGHT_0, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "Wt", 2, "", "",DT_MM41_8022_TABLECELL_WT_1, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "Wt", 3, "", "",DT_MM41_8022_TABLECELL_WT_2, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "Wt", 4, "", "",DT_MM41_8022_TABLECELL_WT_3, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "Length", 1, "", "",DT_MM41_8022_TABLECELL_LENGTH_0, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "Length", 2, "", "",DT_MM41_8022_TABLECELL_LENGTH_1, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "Length", 3, "", "",DT_MM41_8022_TABLECELL_LENGTH_2, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "Length", 4 ,"", "",DT_MM41_8022_TABLECELL_LENGTH_3, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "Width", 1, "", "",DT_MM41_8022_TABLECELL_WIDTH_0, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "Width", 2, "", "",DT_MM41_8022_TABLECELL_WIDTH_1, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "Width", 3, "", "",DT_MM41_8022_TABLECELL_WIDTH_2, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "Width", 4 ,"", "",DT_MM41_8022_TABLECELL_WIDTH_3, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "Height", 1, "", "",DT_MM41_8022_TABLECELL_HEIGHT_0, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "Height", 2, "", "",DT_MM41_8022_TABLECELL_HEIGHT_1, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "Height", 3, "", "",DT_MM41_8022_TABLECELL_HEIGHT_2, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "Height", 4 ,"", "",DT_MM41_8022_TABLECELL_HEIGHT_3, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "Unit of Dimension", 1, "", "",DT_MM41_8022_TABLECELL_UNIT_OF_DIMENSION_0, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "Unit of Dimension", 2, "", "",DT_MM41_8022_TABLECELL_UNIT_OF_DIMENSION_1, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "Unit of Dimension", 3, "", "",DT_MM41_8022_TABLECELL_UNIT_OF_DIMENSION_2, False)
'Call SetTableData("SAPLMGD2TC_ME_8022", "Unit of Dimension", 4 ,"", "",DT_MM41_8022_TABLECELL_UNIT_OF_DIMENSION_3, False)
''Capture the screenshot
'Call TakeScreenShot()
'Call PressEnter() 
'
'Call SetTextbox("ABC Indicator","MAW1-WMAAB","",DT_MM41_2001_ABC_INDICATOR,False)
'Call SetTextbox("Prod\. Hierarchy","MARA-PRDHA",""," ",False)
'Call SetTextbox("Division","MARA-SPART","",DT_MM41_2001_DIVISION,False)
'
''Capture the screenshot
'Call TakeScreenShot()
'
'
''Click on Tax Data
'Call ClickButton("More Tax Data",False)
'Wait(2)
'
'Call SetTableDataNoRef("SAPLMGD2TC_STEUERN","#1",2,DT_MM41_2181_TABLECELL__1,False)
'Call SetTableDataNoRef("SAPLMGD2TC_STEUERN","#3",2,DT_MM41_2181_TABLECELL__1_OCC1,False)
'Call SetTableDataNoRef("SAPLMGD2TC_STEUERN","#5",2,DT_MM41_2181_TABLECELL__1_OCC2,False)
'Call TakeScreenShot()
'
'''Click on Go to main data
'Call ClickButton("Go to main data   \(Ctrl\+Shift\+F3\)",False)
'Wait(2)
'
'Call SetTextbox("Ctry of origin","MAW1-WHERL","",DT_MM41_2002_CTRY_OF_ORIGIN,False)
'Call SetTextbox("Temperature","MARA-TEMPB","",DT_MM41_2002_TEMPERATURE,False)
'Call SetTextbox("Stor\.conditions","MARA-RAUBE","",DT_MM41_2002_STORCONDITIONS,False)
'
'Call SetTextbox("Purch\. Group","MAW1-WEKGR","",DT_MM41_2003_PURCH_GROUP,False)
'Call SetTextbox("Supply source","MARA-BWSCL","",DT_MM41_2003_SUPPLY_SOURCE,False)
'Call SetTextbox("Trans\. Group","MARA-TRAGR","",DT_MM41_2003_TRANS_GROUP,False)
'Call SetTextbox("Loading Group","MAW1-WLADG","",DT_MM41_2003_LOADING_GROUP,False)
'Call TakeScreenShot()
'
'''Click on Next Page
'Call ClickButton("Next Page",False)
'Wait(2)
'
'''Click on Next Page
'Call ClickButton("Next Page",False)
'Wait(2)
'
'Call SetTextbox("Country of Origin","RCTMS-MWERT","",DT_MM41_4000_RCTMSMWERT,False)
'Call SetTextbox("Item strategy","RCTMS-MWERT","",DT_MM41_4000_RCTMSMWERT_OCC1,False)
'Call SetTextbox("Multi-tier assortment pricing","RCTMS-MWERT","",DT_MM41_4000_RCTMSMWERT_OCC2,False)
'Call SetTextbox("Price strategy Maxi","RCTMS-MWERT","",DT_MM41_4000_RCTMSMWERT_OCC3,False)
'Call SetTextbox("Price strategy Tempo","RCTMS-MWERT","",DT_MM41_4000_RCTMSMWERT_OCC4,False)
'
''Enter the details
'Call SetTextbox("Article","RMMW1-MATNR","",DT_MM42_0100_ARTICLE,False)
'Call SetTextbox("Sales Org\.","RMMW1-VKORG","",DT_MM42_0100_SALES_ORG,False)
'Call SetTextbox("Distr\. Channel","RMMW1-VTWEG","",DT_MM42_0100_DISTR_CHANNEL,False)
'
''Capture the screenshot
'Call TakeScreenShot()
'
'Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100","Screen description","Basic Data",False)
'Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100","Screen description","Purchasing",False)
''Capture the screenshot
'Call TakeScreenShot()
'Call PressEnter()
'wait(2)
'Call TakeScreenShot()
'Call PressEnter()
'wait(2)
'
''Navigate to Purchasing Tab
'Call SelectTab("TABSPR1"," Purchasing",False)
'Wait(1)
'Call TakeScreenShot()
'
'Call SetTextbox("Rem\. Shelf Life","EINE-MHDRZ","",DT_MM42_2222_REM_SHELF_LIFE,False)
'Call SetTextbox("Minimum Qty","EINE-MINBM","",DT_MM42_2222_MINIMUM_QTY,False)
'Call SetTextbox("Pl\. Deliv\. Time","EINE-APLFZ","",DT_MM42_2222_PL_DELIV_TIME,False)
''Capture the screenshot
'Call TakeScreenShot()
'
''Save the Details
'Call ClickButton("Save   \(Ctrl\+S\)",False)
'Wait(2)
'Call TakeScreenShot()
'
''Click on Enterif pop up appears
'Call ClickButtonIfExist("Enter   \(F5\)",True)
'Wait(2)
'
''Verify the Message Type
'Call VerifyStatusBarMessageType("S")
'
''Verify the status bar message
'Call VerifyStatusBar(DT_MM42_0100_CHECK_TEXT_OF_STATUSBAR)
'

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus()


'*********************************************End Of Script*********************************************************************









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




