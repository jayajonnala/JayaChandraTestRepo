
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AT_MD_01_01_055-BOM - Display Tabasco     
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


gstrTestCaseName = "Test_AT_MD_01_01_055-BOM - Display Tabasco"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_MD_01_01_055-BOM - Display Tabasco.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


'''--------TransactionCode-MM42----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Article","RMMW1-MATNR","",DT_MM42_0100_ARTICLE,False)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot

Call ClickButtonIfExist("Goto Component Creation",False)

Call SetTableData("SAPLWST1TC_COMPONENTS", "Component", 1, "", "", DT_MM42_0100_TABLECELL_COMPONENT_0, false)
Call SetTableData("SAPLWST1TC_COMPONENTS", "Component qty", 1, "", "", DT_MM42_0100_TABLECELL_COMPONENT_QTY_0, false)
Call PressEnter() 
Call TakeScreenShot

Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot

Call VerifyStatusBar("Article "&DT_MM42_0100_ARTICLE&" changed")

'''--------TransactionCode-MM43----------''''

Call SetTcode(DT_MM42_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
'Call CheckTCodeScreen(DT_MM42_0100_OKCD)

Call PressEnter()     
Call TakeScreenShot

Call ClickButtonIfExist("Goto Component Creation",False)
Call TakeScreenShot

''''Validations
Call VerifyTextBoxContent("MI Test Automation ZHAW", "WSTR_DYNP-DISNR", "", DT_MM42_0103_CHECK_TEXT_OF_MI_TEST_AUTOMATION_ZHAW, False)

Call VerifyTableCellContent(1, "Component", "SAPLWST1TC_COMPONENTS", DT_MM42_0100_CHECK_TEXT_OF_TABLECELL_COMPONENT_0)

Call VerifyTableCellContent(1, "Component qty", "SAPLWST1TC_COMPONENTS", trim(DT_MM42_0100_CHECK_TEXT_OF_TABLECELL_COMPONENT_QTY_0))

Call VerifyTableCellContent(1, "Component unit", "SAPLWST1TC_COMPONENTS", DT_MM42_0100_CHECK_TEXT_OF_TABLECELL_COMPONENT_UNIT_0)


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


