
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_REB_39-ZME3 -  R9 cigarettes
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
'.................Test Script Name :Test_REB_39-ZME3 -  R9 cigarettes
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//




gstrTestCaseName = "Test_REB_39-ZME3 -  R9 cigarettes"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\smasu\Documents\TASE\Data Input\MI\TASE_DT_REB_19-P2P Rebates_ZMU2 - Auto Invoice -R2F-.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)

Call Login(DT_SAPUSER,DT_SAPPASSWORD)

Call PressEnter()

Call SetTcode(DT_SAPTRANSACTIONCODE)     

Call PressEnter()

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call ClickButton("Create Condition Contract   \(F6\)",False)

Call SetComboByKey("KOMWCOCOH-CONTRACT_TYPE", DT_WCOCO_0100_CONDITION_CONTRACT_TYPE)

Call TakeScreenShot()

Call PressEnter()

Call SetTextBox("from","KOMWCOCOH-DATE_FROM",0,ConvertDate(DT_WCOCO_0510_FROM),False)

Call SetTextBox("to","KOMWCOCOH-DATE_TO",0,ConvertDate(DT_WCOCO_0510_TO),False)

Call SetTextBox("Owner Vend","KOMWCOCOH-VEND_OWNER",0,DT_WCOCO_0710_OWNER_VEND,False)

Call SetTextBox("External Number","KOMWCOCOH-EXT_NUM",0,DT_WCOCO_0621_EXTERNAL_NUMBER,False)

Call SetTextBox("Contract Currency","KOMWCOCOH-CC_CURR",0,DT_WCOCO_0621_CONTRACT_CURRENCY,False)

Call PressEnter()

Call TakeScreenShot()

Call SelectTab("TABSTRIP","Purch.",False)

Call SetTextBox("Purchasing Group","KOMWCOCOH-EKGRP",0,DT_WCOCO_0623_PURCHASING_GROUP,False)

Call TakeScreenShot()

Call SelectTab("TABSTRIP","Business volume base",False)

If DT_SAPSYSTEM = "RGB - SAP RETAIL Sandbox GLOBAL" Then
	Call ClickButtonToolBar("WB2R_NEW_LINE", 2)
Else
	Call ClickButtonToolBar("WB2R_NEW_LINE", 1)
End If

Call TakeScreenShot()

Call  SetFocusGuiLabel(DT_WCOCO_0800_FIND, 11, 56, True)

Call PressEnter()

Call TakeScreenShot()

'''Uncomment below step and comment next step with supplier for execution in R1E
'Call SetMultipleGridData("Business Volume Base", 0, 1, "Vendor", DT_WCOCO_0628_GRIDCELL_0_VENDOR, False)
Call SetMultipleGridData("Business Volume Base", 0, 1, "Supplier", DT_WCOCO_0628_GRIDCELL_0_VENDOR, False)

Call SetMultipleGridData("Business Volume Base", 0, 1, "Purchasing Group", DT_WCOCO_0628_GRIDCELL_0_PGR, False)

Call SetMultipleGridData("Business Volume Base", 0, 1, "Company Code", DT_WCOCO_0628_GRIDCELL_0_COCD, False)

Call SetMultipleGridData("Business Volume Base", 0, 1, "Vendor Subrange", DT_WCOCO_0628_GRIDCELL_0_VSR, False)

Call PressEnter()

Call TakeScreenShot()

Call SelectTab("TABSTRIP","Settlement Calendar",False)

Call TakeScreenShot()

If DT_SAPSYSTEM = "RGB - SAP RETAIL Sandbox GLOBAL" Then
	Call ClickButtonToolBar("WB2R_CAL_GENERATE", 2)
Else
	Call ClickButtonToolBar("WB2R_CAL_GENERATE", 1)
End If

Call PressEnter()

Call TakeScreenShot()

If DT_SAPSYSTEM = "RGB - SAP RETAIL Sandbox GLOBAL" Then
	Call ClickButtonToolBar("CO_ADD",1)
Else
	Call ClickButtonToolBar("CO_ADD",0)
End If

Call TakeScreenShot()

'''Uncomment below step and comment next step with supplier for execution in R1E
'Call SetMultipleGridData("Conditions", 0, 1, "Vendor", DT_WCOCO_0510_GRIDCELL_0_VENDOR, False)
Call SetMultipleGridData("Conditions", 0, 1, "Supplier", DT_WCOCO_0510_GRIDCELL_0_VENDOR, False)

Call SetMultipleGridData("Conditions", 0, 1, "Purch. Organization", DT_WCOCO_0510_GRIDCELL_0_PORG, False)

Call SetMultipleGridData("Conditions", 0, 1, "Condition Type", DT_WCOCO_0510_GRIDCELL_0_COND_TYPE, False)

Call SetMultipleGridData("Conditions", 0, 1, "Condition Rate", DT_WCOCO_0510_GRIDCELL_0_CONDITION_RATE, False)

Call SetMultipleGridData("Conditions", 0, 1, "Condition Currency", DT_WCOCO_0510_GRIDCELL_0_UNIT, False)

Call SetMultipleGridData("Conditions", 0, 1, "Valid From", DT_WCOCO_0510_GRIDCELL_0_VALID_FROM, False)

Call SetMultipleGridData("Conditions", 0, 1, "Valid to", DT_WCOCO_0510_GRIDCELL_0_VALID_TO, False)

Call PressEnter()

Call ClickButton("Release Document Header   \(Shift\+F1\)",False)

Call TakeScreenShot()

Call ClickButton("btn\[11\]",False)

Call TakeScreenShot()

Call GetGridContent("Messages", 0, "Line Content", 1, "<NA>", "<NA>", "DT_WCOCO_0300_CHECK_GETCELLVALUE_OF_GRIDCELL_0_TEXT_OUTPUT")

Call LogOff()

Call FinalStatus()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''





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




