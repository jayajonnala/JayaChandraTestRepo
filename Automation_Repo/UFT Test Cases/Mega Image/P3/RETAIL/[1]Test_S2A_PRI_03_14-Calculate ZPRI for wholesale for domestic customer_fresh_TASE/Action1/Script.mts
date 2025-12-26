
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRI_03_14-Calculate ZPRI for wholesale for domestic customer_fresh 
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


gstrTestCaseName = "Test_S2A_PRI_03_14-Calculate ZPRI for wholesale for domestic customer_fresh"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_S2A_PRI_03_06-Recommended price calculation for  different assortment per affiliates_P2 (Fresh).xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''''''-----Login----------'''''''

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''''--------TransactionCode-VKP5----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call ClickButton("Get Variant...   \(Shift\+F5\)",False)
Call SetTextbox("Variant","V-LOW","",DT_VKP5_0100_VARIANT,True)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call ClickButton("Execute   \(F8\)",True)
''Grid title has been changed
''Call SelectRowGuiGridbyRowNo("Variant Catalog for Program RWVKP007", 0, 10, True)
Call SelectRowGuiGridbyRowNo("Variant Catalog.*", 0, 10, True)
Call DoubleClickGuiGridCell("Variant Catalog.*", 0, 10, "Variant name", True)
Call ClickButton("%_S_MATNR_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "<NA>", "<NA>", DT_VKP5_3010_TABLECELL_SINGLE_VALUE_0, False)
Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "<NA>", "<NA>", DT_VKP5_3010_TABLECELL_SINGLE_VALUE_1, False)
Call SetTableData("SAPLALDBSINGLE", "Single value", 3, "<NA>", "<NA>", DT_VKP5_3010_TABLECELL_SINGLE_VALUE_2, False)
Call ClickButton("Copy   \(F8\)",True)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call GetGridContent("", 0, "Basic Pur.Pr.", 1, "<NA>", "<NA>", "DT_GET_1_OUTPUT")
Call GetGridContent("", 0, "Basic Pur.Pr.", 2, "<NA>", "<NA>", "DT_GET_2_OUTPUT")
Call GetGridContent("", 0, "Basic Pur.Pr.", 3, "<NA>", "<NA>", "DT_GET_3_OUTPUT")
Call GetGridContent("", 0, "Purch.net/net", 1, "<NA>", "<NA>", "DT_GET_4_OUTPUT")
Call GetGridContent("", 0, "Purch.net/net", 2, "<NA>", "<NA>", "DT_GET_5_OUTPUT")
Call GetGridContent("", 0, "Purch.net/net", 3, "<NA>", "<NA>", "DT_GET_6_OUTPUT")
Call VerifyGridCellContent("", 1, "Purch.net/net", 0, DT_GET_1_OUTPUT)
Call VerifyGridCellContent("", 2, "Purch.net/net", 0, DT_GET_2_OUTPUT)
Call VerifyGridCellContent("", 3, "Purch.net/net", 0, DT_GET_3_OUTPUT)
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
Call ClickButtonIfExist("Continue   \(Enter\)", True)
Call TakeScreenShot


''''''''--------TransactionCode-VKP6----------''''

Call SetTcode(DT_VKP5_1000_OKCD)     
Call PressEnter() 

Call SetTextbox("Pricing document","P_KBELN","",DT_VKP5_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT,False)
Call PressEnter() 
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "BPSTA", 0, DT_VKP5_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BPSTA)
Call VerifyGridCellContent("", 2, "BPSTA", 0, DT_VKP5_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BPSTA)
Call VerifyGridCellContent("", 3, "BPSTA", 0, DT_VKP5_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BPSTA)

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




