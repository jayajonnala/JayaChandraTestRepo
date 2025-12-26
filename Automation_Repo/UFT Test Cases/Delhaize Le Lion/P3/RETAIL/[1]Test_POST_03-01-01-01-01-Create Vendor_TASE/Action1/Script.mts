
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : 
'.................Test Scenario: Test_POST_03-01-01-01-01-Create Vendor
'.................TCode: XK01
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_POST_03-01-01-01-01-CV"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'ensure no open session
''Call CloseSessionsSAP()
'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''--------TransactionCode-XK02----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SelectCheckbox("RF02K-D0120",1,"ON", False)
Call SetTextboxNoLabel("RF02K-LIFNR","",DT_VENDOR,False)
''This function SetTextboxNoLabel is used for vendor to supplier change.
Call SetTextbox("Company Code","RF02K-BUKRS","",DT_XK02_0101_COMPANY_CODE,False)
Call SetTextbox("Purchasing Organization","RF02K-EKORG","",DT_XK02_0101_PURCH_ORGANIZATION,False)
Call TakeScreenShot
Call PressEnter()
'Call SetTextbox("VAT Reg. No.","LFA1-STCEG","","",False)'Vat No. Issue not yet resolved.
Call ClickButton("Save   \(Ctrl\+S\)", False)
Call PressEnter()
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call TakeScreenShot
'Call GetStatusBar("item2", "DT_OP_EUP_NO") 'Vat No. Issue not yet resolved.
''GetRowNo =4
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call SelectCheckbox("RF02K-D0210",1,"ON", False)
'Call VerifyStatusBar(DT_XK02_0101_CHECK_TEXT_OF_STATUSBAR) 'Vat No. Issue not yet resolved.
Call TakeScreenShot
Call LogOff()
Call FinalStatus ()
'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [13,4062483]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

