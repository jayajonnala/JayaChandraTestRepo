
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_PM002 - Create Maintenance Notification for Echipamente fri 
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


gstrTestCaseName = "Test_PM002 - Create Maintenance Notification for Echipamente fri"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\FICO\TASE_DT_PM002 - Create Maintenance Notification for Echipamente fri.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


''''--------TransactionCode-/niw22----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Notification","RIWO00-QMNUM","",DT_IW22_0100_NOTIFICATION,False)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
Call ClickButton("btn\[13\]",False)
Call TakeScreenShot
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call VerifyStatusBar(DT_IW22_0100_CHECK_TEXT_OF_STATUSBAR)


'''''--------TransactionCode-/niw28----------''''

Call SetTcode(DT_IW22_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(VERIFY_DT_IW22_0100_OKCD)


Call SetTextbox("Notification","QMNUM-LOW","",DT_IW22_0100_NOTIFICATION,False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call VerifyTextBoxContent("Notific. Status", "RIWO00-STTXT", "", DT_IW22_1050_CHECK_TEXT_OF_NOTIFIC_STATUS, False)
Call SelectMenuBar("Extras;Notification documents;Document flow")
Call TakeScreenShot
Call ActivateItemGuiTree(0,"#1",DT_IW22_0100_CHECK_TEXT_OF_ORDER)

 '''''--------TransactionCode---ZMDPM_AUTO_CREATE_WO---------''''

Call SetTcode(DT_IW22_0100_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(VERIFY_DT_IW22_0100_OKCD_OCC1)

Call SetTextbox("Notification","SO_QMNUM-LOW","",DT_IW22_0100_NOTIFICATION,False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

''''--------TransactionCode-/niw28----------''''

Call SetTcode(DT_IW22_0500_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(VERIFY_DT_IW22_0100_OKCD)


Call SetTextbox("Notification","QMNUM-LOW","",DT_IW22_0100_NOTIFICATION,False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call VerifyTextBoxContent("Notific. Status", "RIWO00-STTXT", "", DT_IW22_1050_CHECK_TEXT_OF_NOTIFIC_NOPRORAS, False)
Call SelectMenuBar("Extras;Notification documents;Document flow")
Call TakeScreenShot
Call ActivateItemGuiTree(0,"#1",DT_IW22_0100_CHECK_TEXT_OF_ORDERASSIGNED)
Call ClickButtonIfExist("Back   \(F3\)",False)
Call ActivateNodeGuiTree(0, "#1;#1")
Call GetTextboxValue("CAUFVD-AUFNR", "", "DT_IW22_0100_GET_TEXT_OF_MAINTENANCEORDER_ORDER_OUTPUT", False)
Call ClickButtonIfExist("Back   \(F3\)",False)
Call ActivateNodeGuiTree(0, "#1;#1;#1;#1")
Call GetTextboxValue("MEREQ_TOPLINE-BANFN_EXT", "", "DT_IW22_0100_GET_TEXT_OF_PREQ10_OUTPUT", False)
Call ClickButtonIfExist("Back   \(F3\)",False)
Call ActivateNodeGuiTree(0, "#1;#1;#1;#2")
Call ClickButtonIfExist("Back   \(F3\)",False)
Call ActivateNodeGuiTree(0, "#1;#1;#1;#3")
Call ClickButtonIfExist("Back   \(F3\)",False)
Call SelectMenuBar("Edit;Refresh")
Call ActivateNodeGuiTree(0,"#1;#1;#1;#1;#1")
'***************** Will be reflected in 10 mins*******************'
Call GetTextboxValue("MEPO_TOPLINE-EBELN", "", "DT_IW22_0100_GET_TEXT_OF_PO10_ORDER_OUTPUT", False)
Call ClickButtonIfExist("Back   \(F3\)",False)
Call ActivateNodeGuiTree(0, "#1;#1;#1;#2;#1")
Call ClickButtonIfExist("Back   \(F3\)",False)
Call ActivateNodeGuiTree(0, "#1;#1;#1;#3;#1")
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


