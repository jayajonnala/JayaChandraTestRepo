
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_REB_10-P2P Rebates_ZMU7 Auto Invoice
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
'.................Test Script Name :Test_REB_10-P2P Rebates_ZMU7 Auto Invoice
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_REB_10-P2P Rebates_ZMU7 Auto Invoice"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\smasu\Documents\TASE\Data Input\MI\DT_REB_10-P2P Rebates_ZMU7 Auto Invoice.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)    '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)

Call Login(DT_SAPUSER,DT_SAPPASSWORD)

Call PressEnter()

Call SetTcode(DT_SAPTRANSACTIONCODE)     

Call PressEnter()

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call ClickButton("Other Condition Contract   \(Shift\+F5\)",False)

Call SetTextboxNoLabel("KOMWCOCOH-NUM", 0, DT_WCOCO_0100_CONDITION_CONTRACT, True)

Call PressEnter()

Call SelectRowGuiGrid("Conditions", 0, "Purch. Organization", "RP01", False)

Call ClickButton("Change Condition Contract   \(F7\)",False)

Call SelectRowGuiGrid("Conditions", 0, "Purch. Organization", "RP01", False)


'If DT_SAPSYSTEM = "RGB - SAP RETAIL Sandbox GLOBAL" Then
'	Call ClickButtonToolBar("CO_CHG_DATE",1)
'Else
'	Call ClickButtonToolBar("CO_CHG_DATE",0)
'End If
Call ClickButtonToolBar("CO_CHG_DATE",0)

Call SetTextBox("Valid from","KOMWCOCOI-DATE_FROM",0,ConvertDate(DT_WCOCO_0100_VALID_FROM),True)

Call SetTextBox("Valid to","KOMWCOCOI-DATE_TO",0,ConvertDate(DT_WCOCO_0100_VALID_TO),True)

Call PressEnter()

Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)

Call GetGridContent("Messages", 0, "Line Content", 1, "<NA>", "<NA>", "DT_WCOCO_0300_CHECK_GETCELLVALUE_OF_GRIDCELL_0_TEXT_OUTPUT")

Call VerifyGridCellContent("Messages", 1,"Line Content", 0, DT_WCOCO_0300_CHECK_GETCELLVALUE_OF_GRIDCELL_0_TEXT_OUTPUT)

Call PressEnter()

Call TakeScreenShot()

Call ClickButton("Other Condition Contract   \(Shift\+F5\)",False)

Call PressEnter()

Call TakeScreenShot()

Call SelectMenuBar("Condition Contract;Copy as Successor")

Call SetComboByKey("KOMWCOCOH-CONTRACT_TYPE", DT_WCOCO_0100_CONDITION_CONTRACT_TYPE)

Call PressEnter()

Call SelectRowGuiGrid("Conditions", 0, "Purch. Organization", "RP01", False)

'
'If DT_SAPSYSTEM = "RGB - SAP RETAIL Sandbox GLOBAL" Then
'	Call ClickButtonToolBar("CO_CHG_DATE",1)
'Else
'	Call ClickButtonToolBar("CO_CHG_DATE",0)
'End If
Call ClickButtonToolBar("CO_CHG_DATE",0)

Call SetTextBoxNoLabel("KOMWCOCOI-DATE_FROM",0,ConvertDate(DT_WCOCO_0100_VALID_FROM_OCC1),True)

Call SetTextBox("Valid to","KOMWCOCOI-DATE_TO",0,ConvertDate(DT_WCOCO_0100_VALID_TO_OCC1),True)

Call PressEnter()

Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)

Call GetGridContent("Messages", 0, "Line Content", 1, "<NA>", "<NA>", "DT_WCOCO_0300_CHECK_GETCELLVALUE_OF_GRIDCELL_0_TEXT_OCC2_OUTPUT")

Call VerifyGridCellContent("Messages", 1,"Line Content", 0, DT_WCOCO_0300_CHECK_GETCELLVALUE_OF_GRIDCELL_0_TEXT_OCC2_OUTPUT)



Call TakeScreenShot()

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




