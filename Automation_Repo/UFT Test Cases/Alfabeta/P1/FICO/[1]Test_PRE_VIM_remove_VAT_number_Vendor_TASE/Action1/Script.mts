'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :[1]Test_PRE_VIM_remove_VAT_number_Vendor_TASE
'.................Author : TCS_Ramesh
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


gstrTestCaseName = "[1]Test_PRE_VIM_remove_VAT_number_Vendor_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\rtod\Documents\Input Datasheet\DLL\DT_PRE_VIM_remove_VAT_number_Vendor_TASE1.xls"

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario


Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter()     ' 
'Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call TakeScreenShot()

Call SelectCheckbox("RF02K-D0120",0,"ON",False)
Call SetTextboxNoLabel("RF02K-LIFNR","",DT_XK02_0101_VENDOR,False)
'This function SetTextboxNoLabel is used for vendor to supplier change.
Call PressEnter()

Call GetTextboxValue("LFA1-STCD1","","DT_XK02_0120_CHECK_TEXT_OF_TAX_NUMBER_1_Output",False)
Call GetTextboxValue("LFA1-STCD2","","DT_XK02_0120_CHECK_TEXT_OF_TAX_NUMBER_2_Output",False)
Call GetTextboxValue("LFA1-STCD3","","DT_XK02_0120_CHECK_TEXT_OF_TAX_NUMBER_3_Output",False)
Call GetTextboxValue("LFA1-J_1KFREPRE","","DT_XK02_0120_CHECK_TEXT_OF_REPS_NAME_Output",False)
Call GetTextboxValue("LFA1-FISKU","","DT_XK02_0120_CHECK_TEXT_OF_TAX_OFFICE_Output",False)
Call GetTextboxValue("LFA1-STENR","","DT_XK02_0120_CHECK_TEXT_OF_TAX_NUMBER_Output",False)
Call GetTextboxValue("LFA1-STCD5","","DT_XK02_0120_CHECK_TEXT_OF_TAX_NUMBER_5_Output",False)
Call GetTextboxValue("LFA1-STCEG","","DT_XK02_0120_CHECK_TEXT_OF_VAT_REG_NO_Output",False)

Call SetTextbox("VAT Reg\. No\.","LFA1-STCEG","","",False)
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call PressEnter()
'Call GetStatusBar("item1","DT_KS02_0200_GET_TEXT_OF_CHANGES_Output")
Call  VerifyStatusBar(DT_XK02_0120_CHECK_TEXT_OF_STATUSBAR)

Call LogOff()
Call FinalStatus()



'
'
'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [13,015294]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



