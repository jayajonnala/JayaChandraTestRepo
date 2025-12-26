
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



gstrTestCaseName = "Test_P2P_01_01_053-Return process in store R001_P4_Check_MAP"
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

Call SetTextBox("Job Name","BTCH2170-JOBNAME",0,DT_MIGO_2170_JOB_NAME,False)
Call SetTextBox("User Name","BTCH2170-USERNAME",0,DT_MIGO_2170_USER_NAME,False)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()
Call SetTcode(DT_MIGO_3000_OKCD_OCC4)     
Call PressEnter()     
Call TakeScreenShot()

Call SetTextBox("Material","P_MATNR",0,DT_ZMDIM_MBMAPCHANGES_1000_MATERIAL,False)
Call SetTextBox("Valuation Area","P_BWKEY",0,DT_ZMDIM_MBMAPCHANGES_1000_VALUATION_AREA,False)
Call SetTextBox("Reference Document","P_DOCU",0,DT_ZMDIM_MBMAPCHANGES_1000_REFERENCE_DOCUMENT,False)
Call SetTextBox("Year","P_YEAR",0,DT_ZMDIM_MBMAPCHANGES_1000_YEAR,False)
Call PressEnter()     
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()

Call ClickButton("Set filter   \(Ctrl\+F5\)",False)
Call SelectCellGuiGrid("Column Set",0,0,"Column Name",True)
Call ClickButton("Add Filter Criterion \(F7\)",True)
Call ClickButton("Define Filter Values",True)
Call SetTextbox("Reference Document Number","%%DYN001-LOW","",DT_ZMDIM_MBMAPCHANGES_1000_REFERENCE_DOCUMENT,True)
Call TakeScreenShot
Call ClickButton("Execute   \(Enter\)",True)


Call VerifyGridCellContent("", 1, "Status", 0, DT_ZMDIM_MBMAPCHANGES_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ICON)

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




