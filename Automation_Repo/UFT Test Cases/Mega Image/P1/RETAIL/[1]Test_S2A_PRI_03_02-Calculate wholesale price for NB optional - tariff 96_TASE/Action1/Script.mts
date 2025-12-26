
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRI_03_02-Calculate wholesale price for NB optional - tariff 96
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
'.................Test Script Name :Test_S2A_PRI_03_02-Calculate wholesale price for NB optional - tariff 96
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_S2A_PRI_03_02-NB optional - tariff 96"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_S2A_PRI_03_02-Calculate wholesale price for NB optional - tariff 96-.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)

Call Login(DT_SAPUSER,DT_SAPPASSWORD)

Call PressEnter()

Call SetTcode(DT_SAPTRANSACTIONCODE)     

Call PressEnter()   

Call SelectCheckbox("P_SLOG", 1, DT_VKP5_1000_ISSUE_UPDATE_LOG, False)
Call SetTextbox("Sales organization","S_VKORG-LOW","",DT_VKP5_1000_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution channel","S_VTWEG-LOW","",DT_VKP5_1000_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Purchase Price Determ. Seq.","P_EKERV","",DT_VKP5_1000_PURCHASE_PRICE_DETERM_SEQ,False)
Call SetTextbox("Sales price determination seq.","P_VKERV","",DT_VKP5_1000_SALES_PRICE_DETERMINATION_SEQ,False)
Call SetTextbox("List Group","P_LIGRU","",DT_VKP5_1000_LIST_GROUP,False)
Call SetTextbox("List Variant","P_LIVAR","",DT_VKP5_1000_LIST_VARIANT,False)
Call SetTextbox("Validity","S_DATUM-LOW","",ConvertDate(DT_VKP5_1000_VALIDITY),False)
Call SetTextbox("to","S_DATUM-HIGH","",ConvertDate(DT_VKP5_1000_TO),False)
Call TakeScreenShot
Call SetTextbox("Article","S_MATNR-LOW","","",False)
Call ClickButton("%_S_MATNR_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "<NA>", "<NA>", DT_VKP5_3010_TABLECELL_SINGLE_VALUE_0, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "<NA>", "<NA>", DT_VKP5_3010_TABLECELL_SINGLE_VALUE_1, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 3, "<NA>", "<NA>", DT_VKP5_3010_TABLECELL_SINGLE_VALUE_2, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 4, "<NA>", "<NA>", DT_VKP5_3010_TABLECELL_SINGLE_VALUE_3, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 5, "<NA>", "<NA>", DT_VKP5_3010_TABLECELL_SINGLE_VALUE_4, True)
Call ClickButton("Copy   \(F8\)",True)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call SelectRowGuiGrid("", 0, "Article", DT_VKP5_3010_TABLECELL_SINGLE_VALUE_2, False)
Call ClickButtonToolBar("KALK", 0)
Call TakeScreenShot
''Call VerifyTableCellContent(4, "Amount", "SAPLV69ATCTRL_KONDITIONEN", DT_VKP5_6201_CHECK_TEXT_OF_TABLECELL_AMOUNT_3)
Call ClickButton("Back   \(F3\)",False)
Call SelectAllRowGuiGrid("", 0, False)
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call TakeScreenShot
Call GetStatusBar("item1", "DT_VKP5_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Data saved; pricing document "&DT_VKP5_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" created")

''''''--------TransactionCode-/nVKP6----------''''

Call SetTcode(DT_VKP5_1000_OKCD)     
Call PressEnter()     

Call SetTextbox("Pricing document","P_KBELN","",DT_VKP5_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT,False)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "BPSTA", 0, DT_VKP5_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BPSTA)
Call VerifyGridCellContent("", 2, "BPSTA", 0, DT_VKP5_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BPSTA)
Call VerifyGridCellContent("", 3, "BPSTA", 0, DT_VKP5_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BPSTA)
Call VerifyGridCellContent("", 4, "BPSTA", 0, DT_VKP5_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BPSTA)
Call VerifyGridCellContent("", 5, "BPSTA", 0, DT_VKP5_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BPSTA)


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




