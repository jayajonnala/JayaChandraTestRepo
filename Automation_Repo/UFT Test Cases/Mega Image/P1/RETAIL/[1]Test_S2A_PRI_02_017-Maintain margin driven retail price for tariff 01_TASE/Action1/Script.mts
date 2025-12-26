
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRI_02_017-Maintain margin driven retail price for tariff 01
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
'.................Test Script Name :Test_S2A_PRI_02_017-Maintain margin driven retail price for tariff 01
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_S2A_PRI_02_017-Maintain price for tariff 01"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_S2A_PRI_02_017-Maintain margin driven retail price for tariff 01.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter() 

Call SetTextbox("Article","S_MATNR-LOW","","",False)
Call ClickButton("%_S_MATNR_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "<NA>", "<NA>", DT_VKP5_3010_TABLECELL_SINGLE_VALUE_0, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "<NA>", "<NA>", DT_VKP5_3010_TABLECELL_SINGLE_VALUE_1, True)
Call ClickButton("Copy   \(F8\)",True)
Call SetTextbox("Sales organization","S_VKORG-LOW","",DT_VKP5_1000_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution channel","S_VTWEG-LOW","",DT_VKP5_1000_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Price List","S_PLTYP-LOW","","",False)
Call TakeScreenShot
Call ClickButton("%_S_PLTYP_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "<NA>", "<NA>", DT_VKP5_3010_TABLECELL_SINGLE_VALUE_0_OCC1, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "<NA>", "<NA>", DT_VKP5_3010_TABLECELL_SINGLE_VALUE_1_OCC1, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 3, "<NA>", "<NA>", DT_VKP5_3010_TABLECELL_SINGLE_VALUE_2, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 4, "<NA>", "<NA>", DT_VKP5_3010_TABLECELL_SINGLE_VALUE_3, True)
Call ClickButton("Copy   \(F8\)",True)
Call SelectCheckbox("P_SLOG", 1, DT_VKP5_1000_ISSUE_UPDATE_LOG, False)
Call SetTextbox("Purchase Price Determ. Seq.","P_EKERV","",DT_VKP5_1000_PURCHASE_PRICE_DETERM_SEQ,False)
Call SetTextbox("Sales price determination seq.","P_VKERV","",DT_VKP5_1000_SALES_PRICE_DETERMINATION_SEQ,False)
Call SetTextbox("List Group","P_LIGRU","",DT_VKP5_1000_LIST_GROUP,False)
Call SetTextbox("List Variant","P_LIVAR","",DT_VKP5_1000_LIST_VARIANT,False)
Call TakeScreenShot
Call FocusTextBox("Article","S_MATNR-LOW", False)
Call PressEnter() 
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call SelectColumnGuiGrid("", 0, "Article", False)
Call ClickButtonToolBar("&MB_FILTER", 0)
Call SetTextbox("Article","%%DYN001-LOW","",DT_VKP5_1105_ARTICLE,True)
Call TakeScreenShot
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot
Call GetGridContent("", 0, "ENDPA", 1, "<NA>", "<NA>", "DT_VKP5_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ENDPA_OUTPUT")
Call GetGridContent("", 0, "ENDPA", 2, "<NA>", "<NA>", "DT_VKP5_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ENDPA_OUTPUT")
Call GetGridContent("", 0, "ENDPA", 3, "<NA>", "<NA>", "DT_VKP5_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_2_ENDPA_OUTPUT")
Call GetGridContent("", 0, "ENDPA", 4, "<NA>", "<NA>", "DT_VKP5_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_3_ENDPA_OUTPUT")
Call GetGridContent("", 0, "ENDPA", 5, "<NA>", "<NA>", "DT_VKP5_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_4_ENDPA_OUTPUT")
Call SetGridData("", 1, "ENDPR", DT_VKP5_0100_GRIDCELL_0_FINAL_PRICE, False)
Call SetGridData("", 2, "ENDPR", DT_VKP5_0100_GRIDCELL_1_FINAL_PRICE, False)
Call SetGridData("", 3, "ENDPR", DT_VKP5_0100_GRIDCELL_2_FINAL_PRICE, False)
Call SetGridData("", 4, "ENDPR", DT_VKP5_0100_GRIDCELL_3_FINAL_PRICE, False)
Call SetGridData("", 5, "ENDPR", DT_VKP5_0100_GRIDCELL_4_FINAL_PRICE, False)
Call PressEnter() 
Call TakeScreenShot

Call SelectAllRowGuiGrid("", 0, False)
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)

Call TakeScreenShot
Call GetStatusBar("item1", "DT_VKP5_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Data saved; pricing document "&DT_VKP5_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" created")
Call ClickButtonIfExist("Continue   \(Enter\)", True)

''''''--------TransactionCode-/nVKP6----------''''

Call SetTcode(DT_VKP5_1000_OKCD)     
Call PressEnter()     

Call SetTextbox("Pricing document","P_KBELN","",DT_VKP5_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT,False)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "ENDPR", 0, DT_VKP5_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ENDPR)
Call VerifyGridCellContent("", 2, "ENDPR", 0, DT_VKP5_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ENDPR)
Call VerifyGridCellContent("", 3, "ENDPR", 0, DT_VKP5_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_2_ENDPR)
Call VerifyGridCellContent("", 4, "ENDPR", 0, DT_VKP5_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_3_ENDPR)
Call VerifyGridCellContent("", 5, "ENDPR", 0, DT_VKP5_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_4_ENDPR)


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




