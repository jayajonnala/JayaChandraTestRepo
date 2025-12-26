
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRI_03_14-Calculate ZPRI for wholesale for domestic customer_other 
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



gstrTestCaseName = "Test_S2A_PRI_03_14-Calculate ZPRI for wholesale for domestic customer_other"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_S2A_PRI_03_14-Calculate ZPRI for wholesale for domestic customer_other.xls"
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
''grid title has been changed
''Call SelectRowGuiGridbyRowNo("Variant Catalog for Program RWVKP007", 0, 10, True)
Call SelectRowGuiGridbyRowNo("Variant Catalog.*", 0, 10, True)
Call DoubleClickGuiGridCell("Variant Catalog.*", 0, 10, "Variant name", True)
Call ClickButton("%_S_MATNR_%_APP_%-VALU_PUSH",False)
Call ClickButton("btn\[23\]",True)
Call SetTextbox("Directory","DY_PATH","",DT_VKP5_0200_DIRECTORY,True)
Call SetTextbox("File Name","DY_FILENAME","",DT_VKP5_0200_FILE_NAME,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call ClickButton("Copy   \(F8\)",True)
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
Call GetGridContent("", 0, "EKPNN", 10, "<NA>", "<NA>", "DT_GET_PP_10_OUTPUT")
Call GetGridContent("", 0, "EKPNN", 11, "<NA>", "<NA>", "DT_GET_PP_11_OUTPUT")
Call GetGridContent("", 0, "EKPNN", 12, "<NA>", "<NA>", "DT_GET_PP_12_OUTPUT")
Call GetGridContent("", 0, "EKPNN", 13, "<NA>", "<NA>", "DT_GET_PP_13_OUTPUT")
Call GetGridContent("", 0, "EKPNN", 14, "<NA>", "<NA>", "DT_GET_PP_14_OUTPUT")
Call GetGridContent("", 0, "EKPNN", 15, "<NA>", "<NA>", "DT_GET_PP_15_OUTPUT")
Call GetGridContent("", 0, "EKPNN", 16, "<NA>", "<NA>", "DT_GET_PP_16_OUTPUT")
Call GetGridContent("", 0, "EKPNN", 17, "<NA>", "<NA>", "DT_GET_PP_17_OUTPUT")
Call GetGridContent("", 0, "EKPNN", 18, "<NA>", "<NA>", "DT_GET_PP_18_OUTPUT")
Call GetGridContent("", 0, "EKPNN", 19, "<NA>", "<NA>", "DT_GET_PP_19_OUTPUT")
Call GetGridContent("", 0, "EKPNN", 20, "<NA>", "<NA>", "DT_GET_PP_20_OUTPUT")
Call GetGridContent("", 0, "EKPNN", 21, "<NA>", "<NA>", "DT_GET_PP_21_OUTPUT")
Call GetGridContent("", 0, "EKPNN", 22, "<NA>", "<NA>", "DT_GET_PP_22_OUTPUT")
Call GetGridContent("", 0, "EKPNN", 23, "<NA>", "<NA>", "DT_GET_PP_23_OUTPUT")
Call GetGridContent("", 0, "EKPNN", 24, "<NA>", "<NA>", "DT_GET_PP_24_OUTPUT")
Call GetGridContent("", 0, "EKPNN", 25, "<NA>", "<NA>", "DT_GET_PP_25_OUTPUT")
Call GetGridContent("", 0, "EKPNN", 26, "<NA>", "<NA>", "DT_GET_PP_26_OUTPUT")
Call GetGridContent("", 0, "EKPNN", 27, "<NA>", "<NA>", "DT_GET_PP_27_OUTPUT")
Call GetGridContent("", 0, "EKPNN", 28, "<NA>", "<NA>", "DT_GET_PP_28_OUTPUT")

'Call GetGridContent("", 0, "Basic Pur.Pr.", 2, "<NA>", "<NA>", "DT_GET_2_OUTPUT")
'Call GetGridContent("", 0, "Purch.net/net", 1, "<NA>", "<NA>", "DT_GET_4_OUTPUT")

Call VerifyGridCellContent("", 1, "EKPNN", 0, DT_GET_PP_1_OUTPUT)
Call VerifyGridCellContent("", 2, "EKPNN", 0, DT_GET_PP_2_OUTPUT)
Call VerifyGridCellContent("", 3, "EKPNN", 0, DT_GET_PP_3_OUTPUT)
Call VerifyGridCellContent("", 4, "EKPNN", 0, DT_GET_PP_4_OUTPUT)
Call VerifyGridCellContent("", 5, "EKPNN", 0, DT_GET_PP_5_OUTPUT)
Call VerifyGridCellContent("", 6, "EKPNN", 0, DT_GET_PP_6_OUTPUT)
Call VerifyGridCellContent("", 7, "EKPNN", 0, DT_GET_PP_7_OUTPUT)
Call VerifyGridCellContent("", 8, "EKPNN", 0, DT_GET_PP_8_OUTPUT)
Call VerifyGridCellContent("", 9, "EKPNN", 0, DT_GET_PP_9_OUTPUT)
Call VerifyGridCellContent("", 10, "EKPNN", 0, DT_GET_PP_10_OUTPUT)
Call VerifyGridCellContent("", 11, "EKPNN", 0, DT_GET_PP_11_OUTPUT)
Call VerifyGridCellContent("", 12, "EKPNN", 0, DT_GET_PP_12_OUTPUT)
Call VerifyGridCellContent("", 13, "EKPNN", 0, DT_GET_PP_13_OUTPUT)
Call VerifyGridCellContent("", 14, "EKPNN", 0, DT_GET_PP_14_OUTPUT)
Call VerifyGridCellContent("", 15, "EKPNN", 0, DT_GET_PP_15_OUTPUT)
Call VerifyGridCellContent("", 16, "EKPNN", 0, DT_GET_PP_16_OUTPUT)
Call VerifyGridCellContent("", 17, "EKPNN", 0, DT_GET_PP_17_OUTPUT)
Call VerifyGridCellContent("", 18, "EKPNN", 0, DT_GET_PP_18_OUTPUT)
Call VerifyGridCellContent("", 19, "EKPNN", 0, DT_GET_PP_19_OUTPUT)
Call VerifyGridCellContent("", 20, "EKPNN", 0, DT_GET_PP_20_OUTPUT)
Call VerifyGridCellContent("", 21, "EKPNN", 0, DT_GET_PP_21_OUTPUT)
Call VerifyGridCellContent("", 22, "EKPNN", 0, DT_GET_PP_22_OUTPUT)
Call VerifyGridCellContent("", 23, "EKPNN", 0, DT_GET_PP_23_OUTPUT)
Call VerifyGridCellContent("", 24, "EKPNN", 0, DT_GET_PP_24_OUTPUT)
Call VerifyGridCellContent("", 25, "EKPNN", 0, DT_GET_PP_25_OUTPUT)
Call VerifyGridCellContent("", 26, "EKPNN", 0, DT_GET_PP_26_OUTPUT)
Call VerifyGridCellContent("", 27, "EKPNN", 0, DT_GET_PP_27_OUTPUT)
Call VerifyGridCellContent("", 28, "EKPNN", 0, DT_GET_PP_28_OUTPUT)

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
Call ClickButtonIfExist("Continue   \(Enter\)", True)
Call TakeScreenShot


''''''''--------TransactionCode-VKP6----------''''

Call SetTcode(DT_VKP5_1000_OKCD)     
Call PressEnter() 

Call SetTextbox("Pricing document","P_KBELN","",DT_VKP5_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT,False)
Call PressEnter() 
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "BPSTA", 0, DT_CHECK_1)
Call VerifyGridCellContent("", 2, "BPSTA", 0, DT_CHECK_1)
Call VerifyGridCellContent("", 3, "BPSTA", 0, DT_CHECK_1)
Call VerifyGridCellContent("", 4, "BPSTA", 0, DT_CHECK_1)
Call VerifyGridCellContent("", 5, "BPSTA", 0, DT_CHECK_1)
Call VerifyGridCellContent("", 6, "BPSTA", 0, DT_CHECK_1)
Call VerifyGridCellContent("", 7, "BPSTA", 0, DT_CHECK_1)
Call VerifyGridCellContent("", 8, "BPSTA", 0, DT_CHECK_1)
Call VerifyGridCellContent("", 9, "BPSTA", 0, DT_CHECK_1)
Call VerifyGridCellContent("", 10, "BPSTA", 0, DT_CHECK_1)
Call VerifyGridCellContent("", 11, "BPSTA", 0, DT_CHECK_1)
Call VerifyGridCellContent("", 12, "BPSTA", 0, DT_CHECK_1)
Call VerifyGridCellContent("", 13, "BPSTA", 0, DT_CHECK_1)
Call VerifyGridCellContent("", 14, "BPSTA", 0, DT_CHECK_1)
Call VerifyGridCellContent("", 15, "BPSTA", 0, DT_CHECK_1)
Call VerifyGridCellContent("", 16, "BPSTA", 0, DT_CHECK_1)
Call VerifyGridCellContent("", 17, "BPSTA", 0, DT_CHECK_1)
Call VerifyGridCellContent("", 18, "BPSTA", 0, DT_CHECK_1)
Call VerifyGridCellContent("", 19, "BPSTA", 0, DT_CHECK_1)
Call VerifyGridCellContent("", 20, "BPSTA", 0, DT_CHECK_1)
Call VerifyGridCellContent("", 21, "BPSTA", 0, DT_CHECK_1)
Call VerifyGridCellContent("", 22, "BPSTA", 0, DT_CHECK_1)
Call VerifyGridCellContent("", 23, "BPSTA", 0, DT_CHECK_1)
Call VerifyGridCellContent("", 24, "BPSTA", 0, DT_CHECK_1)
Call VerifyGridCellContent("", 25, "BPSTA", 0, DT_CHECK_1)
Call VerifyGridCellContent("", 26, "BPSTA", 0, DT_CHECK_1)
Call VerifyGridCellContent("", 27, "BPSTA", 0, DT_CHECK_1)
Call VerifyGridCellContent("", 28, "BPSTA", 0, DT_CHECK_1)

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




