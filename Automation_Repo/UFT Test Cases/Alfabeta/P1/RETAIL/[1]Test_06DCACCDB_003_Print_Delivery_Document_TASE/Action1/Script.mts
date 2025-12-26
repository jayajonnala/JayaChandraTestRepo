
If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	GetRowNo= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If


'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_06DCACCDB_003_Print_Delivery_Document
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_06DCACCDB_003_Print_Delivery_Document"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_06DCACCDB_003_Print_Delivery_Document_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'GetRowNo=4
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

'--------------------------------VL09-----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_VL02N_4004_OUTBOUND_DELIVERY,False) 
Call TakeScreenShot()
Call PressEnter() 

Call SelectMenuBar("Extras;Delivery Output;Header")
Call SetTableData("SAPDV70ATC_NAST3","Output Type","2","","",DT_VL02N_0100_TABLECELL_OUTPUT_TYPE_1,False) 
Call TakeScreenShot()
Call PressEnter() 
Call SelectRowGuiTable("SAPDV70ATC_NAST3","Output Type",DT_VL02N_0100_TABLECELL_OUTPUT_TYPE_1,False)
Call ClickButton("Display Means of Communication   \(F2\)",False) 
Call SelectCheckbox("NAST-DIMME",0,DT_VL02N_0101_PRINT_IMMEDIATELY,False)
Call SelectCheckbox("NAST-DELET",0,DT_VL02N_0101_RELEASE_AFTER_OUTPUT,False)
Call SetTextbox("Logical destination","NAST-LDEST","",DT_VL02N_0101_LOGICAL_DESTINATION,False)  
Call PressEnter() 
Call TakeScreenShot()
Call ClickButton("Back   \(F3\)",False) 
Call SelectRowGuiTable("SAPDV70ATC_NAST3","Output Type",DT_VL02N_0100_TABLECELL_OUTPUT_TYPE_1,False)
Call ClickButton("Further data   \(F5\)",False) 


Call SetComboByKey("NAST-VSZTP",DT_VL02N_0102_DISPATCH_TIME)
Call TakeScreenShot()
Call ClickButton("Back   \(F3\)",False) 
Call ClickButton("Back   \(F3\)",False) 
Call ClickButton("Save   \(Ctrl\+S\)",False) 
Call VerifyStatusBar(DT_VL02N_4004_CHECK_TEXT_OF_STATUSBAR)

Call PressEnter() 
Call ClickButton("Delete   \(Shift\+F2\)",False) 
Call ClickButtonIfExist("Yes",True) 
Call ClickButtonIfExist("Delete",True)
Call VerifyStatusBar(DT_VL02N_4004_CHECK_TEXT_OF_STATUSBAR_OCC1)

Call ClickButton("Back   \(F3\)",False) 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)

'--------------------------------VA02-----------------------------

Call SetTcode(DT_VL02N_0100_OKCD) 
Call PressEnter()     
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call SetTextbox("Order","VBAK-VBELN","",DT_VL02N_0102_ORDER,False) 
Call PressEnter() 
Call TakeScreenShot() 
Call PressEnter() 
Call SelectMenuBar("Sales document;Delete")
Call ClickButtonIfExist("Yes",True) 
Call VerifyStatusBar(DT_VL02N_0102_CHECK_TEXT_OF_STATUSBAR)

Call LogOff()
Call FinalStatus ()

'




'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [12,8640008]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//





