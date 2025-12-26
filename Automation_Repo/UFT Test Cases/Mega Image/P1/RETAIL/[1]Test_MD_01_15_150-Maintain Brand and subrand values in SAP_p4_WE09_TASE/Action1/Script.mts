'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_01_15_150-Maintain Brand and subrand values in SAP_p4_WE09
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


gstrTestCaseName = "Test_MD_01_15_150 SAP_p4_WE09"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_MD_01_15_150-Maintain Brand and subrand values in SAP_p4_WE09.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''------------------------------------login------------------------------
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''''''--------TransactionCode-WE09----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

''Call SetTextbox("Created At","CRETIM-LOW","",DT_WE09_1000_CREATED_AT,False)
''Call SetTextbox("Created On","CREDAT-LOW","",ConvertDate(DT_WE09_1000_CREATED_ON),False)
Call SetTextbox("Created On","CREDAT-LOW","",ConvertDate(DT_WE09_1000_CREATED_ON),False)
Call SetTextbox("Logical Message","MESTYP-LOW","",DT_WE09_1000_LOGICAL_MESSAGE,False)

Call SetTextbox("Search in Field \.\.\.","FIELD1_1","",DT_WE09_1000_SEARCH_IN_FIELD_,False)
Call SetTextbox("for Value \.\.\.","VALUE1_1",0,DT_FIELD_VALUE,False)

Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()

Call VerifyStatusBar("IDocs were found")

Call ClickLabel("Time",0,False)
Call ClickButton("Sort in descending order   \(Ctrl\+Shift\+F4\)",False)
 
Call TakeScreenShot()

Call GetLabelContentByRefLabel("IDoc number",0,-48,"DT_IDOC_OUTPUT",False)
Call ClickLabel(DT_IDOC_OUTPUT,0,False)
Call TakeScreenShot()

Call ClickLinkGuiTree("IDoc "& DT_IDOC_OUTPUT&";Data records;ZZMDAM_BRAND_HEAD;ZZMDAM_BRAND_TEXTS","ZZMDAM_BRAND_TEXTS","",False)
'Call ClickLinkGuiTree("IDoc "& DT_IDOC_OUTPUT&";Data records;Segment 000001;Segment 000002","ZZMDAM_BRAND_TEXTS","",False)
Call VerifyTextBoxContent("Current Status", "EDIDC-STATUS", "", DT_WE09_0100_CHECK_TEXT_OF_CURRENT_STATUS, False)

Call VerifyTableCellContent(3, "Fld Cont.", "IDOC_TREE_CONTROLINT_SEG_CONTROL", lcase(DT_WE09_0100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_2))
Call VerifyTableCellContent(4, "Fld Cont.", "IDOC_TREE_CONTROLINT_SEG_CONTROL", DT_WE09_0100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_3)
Call TakeScreenShot()

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



