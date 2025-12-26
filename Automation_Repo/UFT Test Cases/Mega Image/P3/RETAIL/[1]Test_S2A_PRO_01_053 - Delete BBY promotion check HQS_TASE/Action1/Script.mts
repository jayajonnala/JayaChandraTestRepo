
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRO_01_053 - Delete BBY promotion check HQS  
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


gstrTestCaseName = "Test_S2A_PRO_01_053 - Delete BBY promotion check HQS"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_S2A_PRO_01_053 - Delete BBY promotion check HQS.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''-----------------------------login----------------------------

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''-----------------------------WAK2----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()   

Call SetTextBox("Promotion","WAKHD-AKTNR",0,DT_WAK2_1100_PROMOTION,False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()
Call SelectTab("TAXI_TABSTRIP_UEBERSICHT","Bonus Buy",False)
Call TakeScreenShot()

Call SelectRowGuiGrid("Bonus Buys - Overview", 0, "Status", "Activated", False)
Call ClickButton("Bonus Buys Deactivate",False)
Call TakeScreenShot()
Call SelectRowGuiGrid("Bonus Buys - Overview", 0, "Status", "Deactivated", False)
Call ClickButton("Bonus Buy Change",False)
Call TakeScreenShot()

Call SetTextBox("Bonus Expiration","RDM_S_BBY_WA_HEADER_UI-ZZ_DATBI",0,ConvertDate(DT_WAK2_0600_BONUS_EXPIRATION),False)
Call PressEnter()
Call GetTextboxValue("RDM_S_BBY_WA_HEADER_UI-STATUS_TXT", "", "DT_WAK2_0600_CHECK_TEXT_OF_STATUS_OUTPUT", False)
Call VerifyTextBoxContent("Status", "RDM_S_BBY_WA_HEADER_UI-STATUS_TXT", "", lcase(DT_WAK2_0600_CHECK_TEXT_OF_STATUS_OUTPUT), False)
Call ClickButton("Activate Bonus Buy   \(F8\)",False)
Call TakeScreenShot()
Call GetTextboxValue("RDM_S_BBY_WA_HEADER_UI-STATUS_TXT", "", "DT_WAK2_0600_CHECK_TEXT_OF_STATUS_OCC2_OUTPUT", False)
Call VerifyTextBoxContent("Status", "RDM_S_BBY_WA_HEADER_UI-STATUS_TXT", "", lcase(DT_WAK2_0600_CHECK_TEXT_OF_STATUS_OCC2_OUTPUT), False)
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot()
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_WAK2_1100_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Promotion "&DT_WAK2_1100_CHECK_TEXT_OF_STATUSBAR_OUTPUT&" saved")
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




