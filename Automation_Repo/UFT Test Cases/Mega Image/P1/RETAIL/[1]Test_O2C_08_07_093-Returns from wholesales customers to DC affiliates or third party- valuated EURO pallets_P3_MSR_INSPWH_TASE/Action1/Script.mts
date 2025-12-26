
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_O2C_08_07_093-Returns from wholesales customers to DC affiliates or third party- valuated EURO pallets_P3_MSR_INSPWH    
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

gstrTestCaseName = "Test_O2C_08_07_093-Returns valuated EURO pallets_P3_MSR_INSPWH"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_O2C_08_07_093-Returns from wholesales customers to DC affiliates or third party- valuated EURO pallets_P3_MSR_INSPWH.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


''''''--------TransactionCode-MSR_INSPWH----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call SetTextbox("Delivery","INBD2-LOW","",DT_MSR_INSPWH_0110_DELIVERY,False)
Call PressEnter() 
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot
Call SelectTab("TABSTRIP_INSP", "Item", False)
Call SetTextbox("Inspection Code","MSR_S_INSP_UI_ITM-DEC_CODE","",DT_MSR_INSPWH_2010_INSPECTION_CODE,False)
Call SetComboByKey("Follow-Up Activity", DT_MSR_INSPWH_2010_FOLLOWUP_ACTIVITY)
Call SelectCheckbox("MSR_S_INSP_UI_ITM-REFUND_RELEASED", 1, DT_MSR_INSPWH_2010_RELEASE_CREDIT_MEMO_REQUEST, False)
Call PressEnter() 
Call TakeScreenShot
Call PressEnter()
Call ClickButtonIfExist("Save and Confim LFU   \(Ctrl\+F8\)",False)
Call TakeScreenShot
Call VerifyStatusBar(DT_MSR_INSPWH_2000_CHECK_TEXT_OF_STATUSBAR)
Call ClickButtonIfExist("Open Returns Overview   \(Ctrl\+F12\)",False)
Call ClickButtonIfExist("Back   \(F3\)",False)
Call ClickButtonIfExist("Open Returns Overview   \(Ctrl\+F12\)",False)
Call ActivateNodeGuiTree(0,"#1;#6")
Call GetTextboxValue("VBAK-VBELN", "", "DT_MSR_INSPWH_4021_CHECK_TEXT_OF_AB_ARM_CREDIT_REQ_OUTPUT", False)
Call TakeScreenShot
Call ClickButtonIfExist("Back   \(F3\)",False)
Call TakeScreenShot

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



