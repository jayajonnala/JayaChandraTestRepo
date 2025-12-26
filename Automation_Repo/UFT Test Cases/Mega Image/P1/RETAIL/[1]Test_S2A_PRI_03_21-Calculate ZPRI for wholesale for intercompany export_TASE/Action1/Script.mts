
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRI_03_21-Calculate ZPRI for wholesale for intercompany export
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
'.................Test Script Name :Test_S2A_PRI_03_21-Calculate ZPRI for wholesale for intercompany export
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_S2A_PRI_03_21-Calculate ZPRI  export"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_S2A_PRI_03_21-Calculate ZPRI for wholesale for intercompany export.xls"
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
Call SetTableData("SAPLALDBSINGLE", "Single value", 3, "<NA>", "<NA>", DT_VKP5_3010_TABLECELL_SINGLE_VALUE_2, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 4, "<NA>", "<NA>", DT_VKP5_3010_TABLECELL_SINGLE_VALUE_3, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 5, "<NA>", "<NA>", DT_VKP5_3010_TABLECELL_SINGLE_VALUE_4, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 6, "<NA>", "<NA>", DT_VKP5_3010_TABLECELL_SINGLE_VALUE_5, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 7, "<NA>", "<NA>", DT_VKP5_3010_TABLECELL_SINGLE_VALUE_6, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 8, "<NA>", "<NA>", DT_VKP5_3010_TABLECELL_SINGLE_VALUE_7, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 9, "<NA>", "<NA>", DT_VKP5_3010_TABLECELL_SINGLE_VALUE_4_OCC1, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 10, "<NA>", "<NA>", DT_VKP5_3010_TABLECELL_SINGLE_VALUE_5_OCC1, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 11, "<NA>", "<NA>", DT_VKP5_3010_TABLECELL_SINGLE_VALUE_6_OCC1, True)
Call ClickButton("Copy   \(F8\)",True)
Call TakeScreenShot
Call SetTextbox("Sales organization","S_VKORG-LOW","",DT_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution channel","S_VTWEG-LOW","",DT_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Price List","S_PLTYP-LOW","",DT_PRICE_LIST,False)
Call SetTextbox("Purchase Price Determ. Seq.","P_EKERV","",DT_PURCHASE_PRICE,False)
Call SetTextbox("Sales price determination seq.","P_VKERV","",DT_SALES_PRICE,False)
Call SetTextbox("List Group","P_LIGRU","",DT_LIST_GROUP,False)
Call SetTextbox("List Variant","P_LIVAR","",DT_LIST_VARIANT,False)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call GetGridContent("", 0, "EKPNN", 1, "<NA>", "<NA>", "DT_GET_PP_1_OUTPUT")
Call GetGridContent("", 0, "EKPNN", 2, "<NA>", "<NA>", "DT_GET_PP_2_OUTPUT")
Call GetGridContent("", 0, "EKPNN", 3, "<NA>", "<NA>", "DT_GET_PP_3_OUTPUT")
Call GetGridContent("", 0, "EKPNN", 4, "<NA>", "<NA>", "DT_GET_PP_4_OUTPUT")
Call GetGridContent("", 0, "EKPNN", 5, "<NA>", "<NA>", "DT_GET_PP_5_OUTPUT")
Call GetGridContent("", 0, "EKPNN", 6, "<NA>", "<NA>", "DT_GET_PP_6_OUTPUT")
Call GetGridContent("", 0, "EKPNN", 7, "<NA>", "<NA>", "DT_GET_PP_7_OUTPUT")
Call GetGridContent("", 0, "EKPNN", 8, "<NA>", "<NA>", "DT_GET_PP_8_OUTPUT")
Call GetGridContent("", 0, "EKPNN", 9, "<NA>", "<NA>", "DT_GET_PP_9_OUTPUT")
Call GetGridContent("", 0, "EKPNN", 9, "<NA>", "<NA>", "DT_GET_PP_9_OUTPUT")
Call GetGridContent("", 0, "EKPNN", 9, "<NA>", "<NA>", "DT_GET_PP_9_OUTPUT")
Call GetGridContent("", 0, "EKPNN", 10, "<NA>", "<NA>", "DT_GET_PP_10_OUTPUT")
Call GetGridContent("", 0, "EKPNN", 11, "<NA>", "<NA>", "DT_GET_PP_11_OUTPUT")
Call GetGridContent("", 0, "ENDPR", 1, "<NA>", "<NA>", "DT_GET_FP_1_OUTPUT")
Call GetGridContent("", 0, "ENDPR", 2, "<NA>", "<NA>", "DT_GET_FP_2_OUTPUT")
Call GetGridContent("", 0, "ENDPR", 3, "<NA>", "<NA>", "DT_GET_FP_3_OUTPUT")
Call GetGridContent("", 0, "ENDPR", 4, "<NA>", "<NA>", "DT_GET_FP_4_OUTPUT")
Call GetGridContent("", 0, "ENDPR", 5, "<NA>", "<NA>", "DT_GET_FP_5_OUTPUT")
Call GetGridContent("", 0, "ENDPR", 6, "<NA>", "<NA>", "DT_GET_FP_6_OUTPUT")
Call GetGridContent("", 0, "ENDPR", 7, "<NA>", "<NA>", "DT_GET_FP_7_OUTPUT")
Call GetGridContent("", 0, "ENDPR", 8, "<NA>", "<NA>", "DT_GET_FP_8_OUTPUT")
Call GetGridContent("", 0, "ENDPR", 9, "<NA>", "<NA>", "DT_GET_FP_9_OUTPUT")
Call GetGridContent("", 0, "ENDPR", 10, "<NA>", "<NA>", "DT_GET_FP_10_OUTPUT")
Call GetGridContent("", 0, "ENDPR", 11, "<NA>", "<NA>", "DT_GET_FP_11_OUTPUT")
Call VerifyGridCellContent("", 1, "ENDPR", 0, DT_GET_PP_1_OUTPUT)
Call VerifyGridCellContent("", 2, "ENDPR", 0, DT_GET_PP_2_OUTPUT)
Call VerifyGridCellContent("", 3, "ENDPR", 0, DT_GET_PP_3_OUTPUT)
Call VerifyGridCellContent("", 4, "ENDPR", 0, DT_GET_PP_4_OUTPUT)
Call VerifyGridCellContent("", 5, "ENDPR", 0, DT_GET_PP_5_OUTPUT)
Call VerifyGridCellContent("", 6, "ENDPR", 0, DT_GET_PP_6_OUTPUT)
Call VerifyGridCellContent("", 7, "ENDPR", 0, DT_GET_PP_7_OUTPUT)
Call VerifyGridCellContent("", 8, "ENDPR", 0, DT_GET_PP_8_OUTPUT)
Call VerifyGridCellContent("", 9, "ENDPR", 0, DT_GET_PP_9_OUTPUT)
Call VerifyGridCellContent("", 10, "ENDPR", 0, DT_GET_PP_10_OUTPUT)
Call VerifyGridCellContent("", 11, "ENDPR", 0, DT_GET_PP_11_OUTPUT)
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

'''''--------TransactionCode-/nVKP6----------''''

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




