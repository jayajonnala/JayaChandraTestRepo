
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRI_01_038-ZCC1 calculated for PO_P3_Condition    
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


gstrTestCaseName = "Test_S2A_PRI_01_038-ZCC1 calculated for PO_P3_Condition"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_S2A_PRI_01_038-ZCC1 calculated for PO_P3_Condition.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''''''-----Login----------'''''''

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''''--------TransactionCode-MEK1----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Condition Type","RV13A-KSCHL","",DT_MEK1_0100_CONDITION_TYPE,False)
Call ClickButton("Condition Information   \(Shift\+F4\)",False)
Call SetTextbox("Purch. Organization","F002-LOW","",DT_MEK1_1000_PURCH_ORGANIZATION,False)
''Call SetTextbox("Vendor","F001-LOW","",DT_MEK1_1000_VENDOR,False)
Call SetTextboxNoLabel("F001-LOW","",DT_MEK1_1000_VENDOR,False)
Call SetTextbox("Article","F003-LOW","",DT_MEK1_1000_ARTICLE,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call SelectCheckboxNoLabel(0, DT_MEK1_0120_NO_NAME, False)
Call ClickButton("Display   \(F5\)",False)
Call TakeScreenShot
Call GetTableCellData("SAPMV13ATCTRL_FAST_ENTRY", "Amount", 1, "<NA>", "<NA>", " DT_MEK1_1084_CHECK_TEXT_OF_TABLECELL_AMOUNT_0_OUTPUT", False)

'''''''--------TransactionCode-MB5S----------''''

Call SetTcode(DT_MEK1_1084_OKCD)     
Call PressEnter()     

Call SetTextbox("Purchasing document","EBELN-LOW","",DT_MEK1_1000_PURCHASING_DOCUMENT,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call SetFocusGuiLabel(DT_MEK1_1000_PURCHASING_DOCUMENT, 11, 88, False)
Call SendKey("{F2}")
Call VerifyGridCellContent("", 1, "BEWTK", 0, DT_MEK1_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BEWTK)
Call VerifyGridCellContent("", 3, "BEWTK", 0, DT_MEK1_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BEWTK)
Call VerifyGridCellContent("", 5, "BEWTK", 0, DT_MEK1_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BEWTK)

'''''''--------TransactionCode-ZMDPC_UPLOAD_COND----------''''

Call SetTcode(DT_MEK1_0400_OKCD)     
Call PressEnter() 

Call SetTextbox("Entry Date","S_CPUDT-LOW","",ConvertDate(DT_MEK1_1000_ENTRY_DATE),False)
'''Call SetTextbox("to","S_CPUDT-HIGH","",ConvertDate((DT_MEK1_1000_ENTRY_DATE)+4),False)
Call SetTextbox("to","S_CPUDT-HIGH","",DT_MEK1_1000_TO,False)
Call SetTextbox("Site","S_WERKS-LOW","",DT_MEK1_1000_SITE,False)
Call SetTextbox("Movement type","S_BWART-LOW","",DT_MEK1_1000_MOVEMENT_TYPE,False)
Call SetTextbox("Purchase order","S_EBELN-LOW","",DT_MEK1_1000_PURCHASING_DOCUMENT,False)
Call TakeScreenShot
Call SelectCheckbox("P_DISP", 0, DT_MEK1_1000_ONLY_DISPLAY, False)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call VerifyifGuiLabelExists(DT_MEK1_0120_CHECK_TEXT_OF_NO_NAME)
Call SelectCheckboxNoLabel(0, DT_MEK1_0120_NO_NAME_OCC1, False)
Call ClickButton("Create Conditions   \(F8\)",False)
Call VerifyStatusBar(DT_MEK1_0120_CHECK_TEXT_OF_STATUSBAR)

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




