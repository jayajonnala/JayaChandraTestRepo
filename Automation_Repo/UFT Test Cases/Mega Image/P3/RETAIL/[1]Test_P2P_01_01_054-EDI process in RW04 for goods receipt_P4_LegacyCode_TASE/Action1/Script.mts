
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_054-EDI process in RW04 for goods receipt_P4_LegacyCode    
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

gstrTestCaseName = "Test_P2P_01_01_054_P4_LegacyCode"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_P2P_01_01_054-EDI process in RW04 for goods receipt_P4_LegacyCode.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''
'
'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


''''''''TransactionCode--ZACTF_LEG_CODES''''''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot


Call SetTextbox("Group Name","P_GROUP","",DT_ZACTF_LEG_CODES_1000_GROUP_NAME,false)
Call SelectRadioButton("P_VEND", "Vendors", false)
Call SetTextbox("Purch\. Organization","P_EKORG","",DT_ZACTF_LEG_CODES_1000_PURCH_ORGANIZATION,false)
''Call SetTextbox("Vendor","S_LIFNR-LOW","",DT_ZACTF_LEG_CODES_1000_VENDOR,false)
Call SetTextboxNoLabel("S_LIFNR-LOW","",DT_ZACTF_LEG_CODES_1000_VENDOR,false)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",false)
Call TakeScreenShot

Call GetGridContent("", "", "Legacy Value", "", "Name 1", "HEIDI CHOCOLAT SA", "DT_ZACTF_LEG_CODES_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LEGACY_VALUE")
Call ClickButton("Back   \(F3\)",False)

Call SelectRadioButton("P_ART", "Articles", false)
Call SetTextbox("Article","S_MATNR-LOW","",DT_ZACTF_LEG_CODES_1000_ARTICLE,false)
Call TakeScreenShot

Call ClickButton("Execute   \(F8\)",false)
Call TakeScreenShot

Call GetGridContent("", "", "Legacy Value", "", "Base Unit of Measure", "EA", "DT_ZACTF_LEG_CODES_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LEGACY_VALUE_OCC1")
Call ClickButton("Back   \(F3\)",False)

Call SetTextbox("Article","S_MATNR-LOW","",DT_ZACTF_LEG_CODES_1000_ARTICLE_OCC1,false)
Call TakeScreenShot

Call ClickButton("Execute   \(F8\)",false)
Call TakeScreenShot

Call GetGridContent("", "", "Legacy Value", "", "Base Unit of Measure", "EA", "DT_ZACTF_LEG_CODES_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LEGACY_VALUE_OCC2")

'//SetTransactionCode- MIRO//

Call SetTcode(DT_ZACTF_LEG_CODES_0500_OKCD)     
Call PressEnter()     

Call SetComboByKey("RM08M-VORGANG", DT_ZACTF_LEG_CODES_6000_TRANSACTION)
''Call SetTextbox("Invoice date","INVFO-BLDAT","",ConvertDate(DT_ZACTF_LEG_CODES_0010_DOCUMENT_DATE),False)
Call SetTextboxNoLabel("INVFO-BLDAT","",ConvertDate(DT_ZACTF_LEG_CODES_0010_DOCUMENT_DATE),False)
Call PressEnter()
Call SetTextbox("Reference","INVFO-XBLNR","",DT_ZACTF_LEG_CODES_0010_REFERENCE,False)
Call PressEnter()

Call SetTextboxNoLabel("RM08M-EBELN", "", DT_ZACTF_LEG_CODES_6211_RM08MEBELN, False)

Call SetComboByKey("RM08M-REFERENZBELEGTYP", DT_ZACTF_LEG_CODES_6211_RM08MXWARE_BNK)
Call TakeScreenShot


Call GetTextboxValue("RM08M-DIFFERENZ", "", "DT_ZACTF_LEG_CODES_6000_CHECK_TEXT_OF_BALANCE_OUTPUT", false)
Call GetTableCellData("SAPLMR1MTC_MR1M", "Quantity", "1", "", "", "DT_CHECK_TEXT_OF_TABLECELL_QUANTITY_0_OUTPUT", False)
Call GetTableCellData("SAPLMR1MTC_MR1M", "Quantity", "2", "", "", "DT_CHECK_TEXT_OF_TABLECELL_QUANTITY_1_OUTPUT", False)

Call GetTableCellData("SAPLMR1MTC_MR1M","Amount","1","","", "DT_AMOUNT_LINE_1_OUTPUT", False)
Call GetTableCellData("SAPLMR1MTC_MR1M","Amount","2","","", "DT_AMOUNT_LINE_2_OUTPUT", False)

Call SelectCheckbox("INVFO-XMWST", 0, DT_ZACTF_LEG_CODES_0010_CALCULATE_TAX, False)

Call GetTextboxValue("RM08M-DIFFERENZ", "", "DT_ZACTF_LEG_CODES_6000_CHECK_TEXT_OF_BALANCE_OCC1_OUTPUT", false)
Call TakeScreenShot

Call ClickButton("Deselect All",false)


Call SelectRowGuiTable("SAPLMR1MTC_MR1M", "Item", DT_ITEM_1, false)
Call SelectCheckbox("INVFO-XMWST", 0, DT_ZACTF_LEG_CODES_0010_CALCULATE_TAX_OFF, False)
Call SelectCheckbox("INVFO-XMWST", 0, DT_ZACTF_LEG_CODES_0010_CALCULATE_TAX, False)
Call GetTextboxValue("INVFO-WMWST", "", "DT_GET_TAX_AMOUNT_1_OUTPUT", false)

Call ClickButton("Deselect All",false)

Call SelectRowGuiTable("SAPLMR1MTC_MR1M", "Item", DT_ITEM_2, false)
Call SelectCheckbox("INVFO-XMWST", 0, DT_ZACTF_LEG_CODES_0010_CALCULATE_TAX_OFF, False)
Call SelectCheckbox("INVFO-XMWST", 0, DT_ZACTF_LEG_CODES_0010_CALCULATE_TAX, False)
Call GetTextboxValue("INVFO-WMWST", "", "DT_GET_TAX_AMOUNT_2_OUTPUT", false)

Call LogOff()

Call FinalStatus ()







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



