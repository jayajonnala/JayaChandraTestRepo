
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_01_01_060-Copy Purchasing Info Records - transfer inactive article to new vendor
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
'.................Test Script Name :[1]Test_P2P_01_01_0279-Duplicate invoice check_P1
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_P2P_01_01_053-R001_P2_Display_Print_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\smasu\Documents\TASE\Data Input\MI\TASE_DT_P2P_01_01_0279-Duplicate invoice check_P1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)

Call Login(DT_SAPUSER,DT_SAPPASSWORD)

Call PressEnter()

Call SetTcode(DT_SAPTRANSACTIONCODE)     

Call PressEnter()     

Call TakeScreenShot()

Call SetComboByKey("GODYNPRO-ACTION",DT_MIGO_0010_GODYNPROACTION)

Call SetComboByKey("GODYNPRO-REFDOC",DT_DOC_TYPE)

Call TakeScreenShot()

Call SetTextboxNoLabel("GODYNPRO-MAT_DOC", 0, DT_MIGO_2010_GODYNPROMAT_DOC, False)

Call PressEnter()

Call SelectTab("TS_GOITEM","Where",False)

Call TakeScreenShot()

Call SelectTab("TS_GOHEAD","Doc. info",False)

Call ClickButton("OK_FI_DOC",False)

Call VerifyGridCellContent("", 1, "Posting Key", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)

Call VerifyGridCellContent("", 2, "Posting Key", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

Call VerifyGridCellContent("", 1, "Account", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)

Call VerifyGridCellContent("", 2, "Account", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)

Call ClickButton("Back   \(F3\)",False)

Call SelectTab("TS_GOITEM","Output",False)

Call TakeScreenShot()

Call ClickButton("OK_NAST_SHOW",False)

Call TakeScreenShot()

Call SetTcode(DT_MIGO_0100_OKCD)

Call PressEnter()

Call TakeScreenShot()

Call SetTextBox("Article Document","RG_MBLNR-LOW",0,DT_MIGO_1000_ARTICLE_DOCUMENT,False)

Call SetTextBox("Article Doc\. Year","PM_MJAHR",0,DT_MIGO_1000_ARTICLE_DOC_YEAR,False)


Call TakeScreenShot()

Call PressEnter()

Call ClickButton("Execute   \(F8\)",False)

Call SelectCheckboxNoLabel(0, "ON", False)

Call ClickButton("Print preview   \(Shift\+F4\)",False)

Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False)

Call ClickButton("Process   \(Shift\+F2\)",False)

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




