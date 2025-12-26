
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_01_01_073-Inactivate article for vendor - ME15 - inactivate info record from artemis - only_p2_WE09
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



gstrTestCaseName = "Test_MD_01_01_073ME15 only_p2_WE09"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_MD_01_01_073-Inactivate article for vendor_WE09.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''--------TransactionCode-WE09----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter()
call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


Call SetTextbox("Created On","CREDAT-LOW","",ConvertDate(DT_WE09_1000_CREATED_ON),False)
Call SetTextbox("Logical Message","MESTYP-LOW","",DT_WE09_1000_LOGICAL_MESSAGE,False)

Call SetTextbox("Search in Field \.\.\.","FIELD1_1","",DT_WE09_1000_SEARCH_IN_FIELD_,False)

Call SetTextbox("for Value \.\.\.","VALUE1_1",0,DT_WE09_1000_FOR_VALUE_,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call VerifyStatusBar("IDocs were found")
''Call VerifyifGuiLabelExists("Data passed.*")
''Call GetLabelContentByRefLabel("Data passed.*",497,0,"DT_IDOC_OUTPUT",False)
'Call GetLabelContentByRefLabel("IDoc number",0,-80,"DT_IDOC_OUTPUT",False)
Call GetLabelContentByRefLabel("IDoc number",0,-48,"DT_IDOC_OUTPUT",False)
Call ClickLabel(DT_IDOC_OUTPUT,0,False)
Call ActivateNodeGuiTree(0, "#1;#2;#1")
Call ClickButton("Back   \(F3\)",False)
Call VerifyTextBoxContent("Current Status", "EDIDC-STATUS", 0, DT_WE09_0100_CHECK_TEXT_OF_CURRENT_STATUS, False)
Call ActivateNodeGuiTree(0, "#1;#2;#1;#1")
Call ClickButton("Back   \(F3\)",False)
Call VerifyTableCellContentByRef("IDOC_TREE_CONTROLINT_SEG_CONTROL","Fld Name","ERDAT","Fld Cont.",DT_WE09_0100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_4)
Call VerifyTableCellContentByRef("IDOC_TREE_CONTROLINT_SEG_CONTROL","Fld Name","MSGFN","Fld Cont.",DT_WE09_0100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_0_OCC1)
Call VerifyTableCellContentByRef("IDOC_TREE_CONTROLINT_SEG_CONTROL","Fld Name","UEBTK","Fld Cont.",DT_WE09_0100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_3)

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



