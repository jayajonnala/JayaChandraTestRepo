
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_O2C_08_07_070-Store return process to DC Claim surplus_P5_MSR_INSPW & Final checks      
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


gstrTestCaseName = "Test_O2C_08_07_070_P5_MSR_INSPW & Final checks"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_O2C_08_07_070-Store return process to DC Claim surplus_P5_MSR_INSPW & Final checks.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


'--------TransactionCode-MSR_INSPWH----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Delivery","INBD2-LOW","",DT_MSR_INSPWH_0110_DELIVERY,True)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",True)
Call SetTextbox("Inspection Code","MSR_S_INSP_UI_HDR-DEC_CODE","",DT_MSR_INSPWH_2001_INSPECTION_CODE,False)
Call SetComboByKey("MSR_S_INSP_UI_HDR-FU_CODE", DT_MSR_INSPWH_2001_FOLLOWUP_ACTIVITY)
Call PressEnter()     
Call TakeScreenShot
Call SelectTab("TABSTRIP_INSP", "Item", False)
Call TakeScreenShot
Call ClickButton("Next Entry",False)
Call TakeScreenShot
Call ClickButton("Save and Confim LFU   \(Ctrl\+F8\)",False)
Call VerifyStatusBar(DT_MSR_INSPWH_2000_CHECK_TEXT_OF_STATUSBAR)
Call ClickButton("Open Returns Overview   \(Ctrl\+F12\)",False)

'''Call VerifyTextBoxContent("Processing Status", "MSR_S_SCR_HEAD_DLV-PROC_STATUS_ICON", "", Lcase(DT_MSR_INSPWH_0213_CHECK_TOOLTIP_OF_PROCESSING_STATUS), False)
'
Call VerifyTextBoxContent("Processing Status", "MSR_S_SCR_HEAD_DLV-PROC_STATUS_ICON", "", DT_MSR_INSPWH_0213_CHECK_TOOLTIP_OF_PROCESSING_STATUS, False)
'Call GetGridContent("", "", "Processing Status", 1, "<NA>", "<NA>", "DT_MSR_INSPWH_0213_CHECK_TOOLTIP_OF_PROCESSING_STATUS_OUTPUT")
Call VerifyGridCellContent("", 1, "Processing Status", "", Lcase(DT_STATUS_ICON))
Call VerifyGridCellContent("", 2, "Processing Status", "", Lcase(DT_STATUS_ICON_OCC1))


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



