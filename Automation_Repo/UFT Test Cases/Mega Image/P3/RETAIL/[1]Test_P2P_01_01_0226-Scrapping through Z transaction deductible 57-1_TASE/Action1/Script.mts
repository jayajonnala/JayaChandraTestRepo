
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_0226-Scrapping through Z transaction deductible      
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


gstrTestCaseName = "Test_P2P_01_01_0226-Z transaction deductible"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_P2P_01_01_0226-Scrapping through Z transaction deductible.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


''''''--------TransactionCode-ZMDIM_DESTR_STOCK----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Location","P_ABLAD", "", DT_ZMDIM_DESTR_STOCK_1000_LOCATION, False)
Call SetTextbox("Reason for Movement","P_GRUND", "", DT_ZMDIM_DESTR_STOCK_1000_REASON_FOR_MOVEMENT, False)
Call SetTextbox("Customer","P_KUNNR", "", DT_ZMDIM_DESTR_STOCK_1000_CUSTOMER, False)
Call SetTextbox("Company Code","S_BUKRS-LOW", "", DT_ZMDIM_DESTR_STOCK_1000_COMPANY_CODE, False)
Call SetTextbox("Site","S_WERKS-LOW", "", DT_ZMDIM_DESTR_STOCK_1000_SITE, False)
Call SetTextbox("Storage Location","P_LGORT", "", DT_ZMDIM_DESTR_STOCK_1000_STORAGE_LOCATION, False)
Call TakeScreenShot

Call ClickButton("%_S_MATNR_%_APP_%-VALU_PUSH", False)
Call TakeScreenShot
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "<NA>", "<NA>", DT_ZMDIM_DESTR_STOCK_3010_TABLECELL_SINGLE_VALUE_0, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "<NA>", "<NA>", DT_ZMDIM_DESTR_STOCK_3010_TABLECELL_SINGLE_VALUE_1, True)
Call TakeScreenShot
Call ClickButton("Copy   \(F8\)", True)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)", False)
Call TakeScreenShot

Call SetGridData("", 1, "MENGE", "4", False)
Call SetGridData("", 2, "MENGE", "5", False)
Call TakeScreenShot
Call DoubleClickGuiGridCell("", 0, 1, "MENGE", False)
Call TakeScreenShot
Call DoubleClickGuiGridCell("", 0, 2, "MENGE", False)
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)", False)
Call TakeScreenShot
Call SelectAllRowGuiGrid("", 0, False)
Call TakeScreenShot
'Call SelectRowRangeGuiGrid("", 0, 1, 2, False)
Call ClickButton("Save Documents   \(F8\)", False)
Call TakeScreenShot
Call GetGridContent("", 0, "MBLNR", 2, "<NA>", "<NA>", "DOC_NAME_OUTPUT")
Call ClickCellGuiGrid("", 0, "MBLNR", 2, "<NA>", "<NA>", False)
Call TakeScreenShot
'''Call VerifyTextBoxNoLabelContent("MSEG-BWART", "", DT_ZMDIM_DESTR_STOCK_0420_CHECK_TEXT_OF_MVT, False)
'''Call VerifyTextBoxNoLabelContent("MSEG-BWART", "", DT_ZMDIM_DESTR_STOCK_0420_CHECK_TEXT_OF_MVT_OCC1, False)
' VerifyGridCellContent(gridTitle, gridRowNumber, gridColumnName, gridIndex, expectedValue)
' VerifyTableCellContent(tableRowNumber, tableColumnName, tableName, expectedValue)
Call VerifyTableCellContent(1,"Movement type","SAPLMIGOTV_GOITEM",DT_ZMDIM_DESTR_STOCK_0420_CHECK_TEXT_OF_MVT)
Call VerifyTableCellContent(2,"Movement type","SAPLMIGOTV_GOITEM",DT_ZMDIM_DESTR_STOCK_0420_CHECK_TEXT_OF_MVT_OCC1)





'''''''--------TransactionCode-/nMB90----------''''

Call SetTcode(DT_ZMDIM_DESTR_STOCK_0420_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call ClickButton("Execute   \(F8\)", False)
Call TakeScreenShot

Call SelectCheckboxNoLabel(0, DT_ZMDIM_DESTR_STOCK_0120_NO_NAME, False)
Call TakeScreenShot
Call ClickButton("Print preview   \(Shift\+F4\)", False)
Call TakeScreenShot

Call SelectMenuBar("Goto;List Display")
Call TakeScreenShot
Call VerifyifGuiLabelExists(lcase(DT_ZMDIM_DESTR_STOCK_0120_CHECK_TEXT_OF_PROCES__VERBAL))
Call VerifyifGuiLabelExists(trim(DT_ZMDIM_DESTR_STOCK_0120_CHECK_TEXT_OF_NO_NAME))
Call VerifyifGuiLabelExists(trim(DT_ZMDIM_DESTR_STOCK_0120_CHECK_TEXT_OF_NO_NAME_OCC1))

'''''''--------TransactionCode-/nMB03----------''''

Call SetTcode(DT_ZMDIM_DESTR_STOCK_0120_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call ClickButton("Overview   \(F5\)", False)
Call TakeScreenShot
Call FocusTextBoxWithIndex("Material Description","MSEG-MATNR", 0, False)
Call TakeScreenShot
Call SendKey("{F2}")
Call TakeScreenShot
Call ClickButton("Output   \(Shift\+F2\)", False)
Call TakeScreenShot
Call SelectRowGuiTableByRow("SAPDV70ATC_NAST3", 1, False)

'''''''--------TransactionCode-/nMB51----------''''

Call SetTcode(DT_ZMDIM_DESTR_STOCK_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call ClickButton("Execute   \(F8\)", False)
Call TakeScreenShot
Call ClickButtonIfExist("Detail List   \(Ctrl\+Shift\+F12\)",False)
Call SelectRowGuiGrid("", 0, "Article Document", DOC_NAME_OUTPUT, False)
Call TakeScreenShot
Call DoubleClickGuiGridCell("", 0, 1, "Article Document", False)
Call TakeScreenShot
Call SelectTab("TS_GOHEAD", "Doc. info", False)
Call TakeScreenShot
Call ClickButton("FI Documents", False)
Call TakeScreenShot
Call GetGridContent("Documents in Accounting", 0, "Document Number", 1, "<NA>", "<NA>", "DT_ZMDIM_DESTR_STOCK_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DOCNR_OUTPUT")
Call ClickButton("Cancel   \(F12\)", True)

''''''--------TransactionCode-/nFB03----------''''

Call SetTcode(DT_ZMDIM_DESTR_STOCK_0001_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Document Number","RF05L-BELNR","",DT_ZMDIM_DESTR_STOCK_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DOCNR_OUTPUT,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_ZMDIM_DESTR_STOCK_0100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_ZMDIM_DESTR_STOCK_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_ZMDIM_DESTR_STOCK_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_ZMDIM_DESTR_STOCK_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_ZMDIM_DESTR_STOCK_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)

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


