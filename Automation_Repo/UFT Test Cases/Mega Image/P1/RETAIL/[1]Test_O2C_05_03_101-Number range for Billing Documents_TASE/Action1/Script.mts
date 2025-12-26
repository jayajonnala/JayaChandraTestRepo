
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_O2C_05_03_101-Number range for Billing Documents
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

gstrTestCaseName = "Test_O2C_05_03_101-Number range for Billing Documents"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_O2C_05_03_101-Number range for Billing Documents.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'GetRowNo = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''-----------------Login-----------------------''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''''''--------TransactionCode-FB03----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call ClickButtonIfExist("Create Document List\/Find Documents   \(Shift\+F8\)",False)

Call SetTextbox("Company code","BR_BUKRS-LOW","",DT_FB03_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","BR_GJAHR-LOW","",DT_FB03_1000_FISCAL_YEAR,False)
Call SetTextbox("Document type","BR_BLART-LOW","",DT_FB03_1000_DOCUMENT_TYPE,False)
Call TakeScreenShot
Call ClickButtonIfExist("%_BR_BLART_%_APP_%-VALU_PUSH",False)
Call TakeScreenShot

Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "", "", DT_FB03_3010_TABLECELL_SINGLE_VALUE_1, True)
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call SetTextbox("Posting date","BR_BUDAT-LOW","",ConvertDate(DT_FB03_1000_POSTING_DATE),False)
Call SetTextbox("to","BR_BUDAT-HIGH","",DT_FB03_1000_TO_OCC1,False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call ClickButtonIfExist("Change Layout\.\.\.   \(Ctrl\+F8\)",False)
'''Call SelectRowGuiGrid("Column Set", "", "Column Name", "Alt Reference Number", True)

Call ClickButtonToolBar("&FIND",0)
Call SetTextboxNoLabel("GS_SEARCH-VALUE","","Alt Reference Number",True)
Call PressEnter()
Call ClickButtonIfExist("Cancel   \(F12\)",True)

Call TakeScreenShot
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)
Call ClickButtonIfExist("Transfer   \(Enter\)",True)
Call TakeScreenShot

Call ClickButtonIfExist("Sort in Ascending Order   \(Ctrl\+F4\)",False)
Call ClickButtonIfExist("Transfer   \(Enter\)",True)
Call TakeScreenShot

Call GetGridContent("", "", "Alt Reference Number", 1, "<NA>", "", "DT_FB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_XBLNR_ALT_OCC1_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_XBLNR_ALT_OCC1_OUTPUT",DT_FB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_XBLNR_ALT_OCC1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call GetGridContent("", "", "Alt Reference Number", 2, "<NA>", "", "DT_FB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_XBLNR_ALT_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_XBLNR_ALT_OUTPUT",DT_FB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_XBLNR_ALT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call GetGridContent("", "", "Alt Reference Number", 3, "<NA>", "", "DT_FB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_XBLNR_ALT_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_XBLNR_ALT_OUTPUT",DT_FB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_XBLNR_ALT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call VerifyGridCellContent("", 1, "Alt Reference Number", "", DT_FB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_XBLNR_ALT)
Call VerifyGridCellContent("", 2, "Alt Reference Number", "", DT_FB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_XBLNR_ALT)
Call VerifyGridCellContent("", 3, "Alt Reference Number", "", DT_FB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_XBLNR_ALT)


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


