

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

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_01PRO11_004_ChangeBonus_ExpDate_CurrentDate_Non_Supported_Scen05
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_01PRO11_004_ChangeBonus_ExpDate_CurrentDate_Non_Supported_Scen05"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRO11_004_ChangeBonus_ExpDate_CurrentDate_Non_Supported_Scen05_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
'
'
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

'--------------------------------------------  WAK2------------------------------------------------

Call SetTextbox("Promotion","WAKHD-AKTNR","",DT_WAK2_1100_PROMOTION,False)
Call TakeScreenShot()
Call  ClickButton("Fast entry   \(F5\)",False)

Call SelectTab("TAXI_TABSTRIP_UEBERSICHT","Bonus Buy",False)
Call SelectRowGuiGrid("Bonus Buys - Overview",0,"Status","Activated",False)
Call  ClickButton("Bonus Buys Deactivate",False)
Call TakeScreenShot()

Call SelectRowGuiGrid("Bonus Buys - Overview",0,"Status","Deactivated",False)
Call  ClickButton("Bonus Buy Change",False)
Call TakeScreenShot()

Call SetTextbox("Bonus Expiration","RDM_S_BBY_WA_HEADER_UI-ZZ_DATBI","",ConvertDate(DT_WAK2_0600_BONUS_EXPIRATION),False)
Call PressEnter()
'Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
'Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC2)

Call SetTextbox("Bonus Expiration","RDM_S_BBY_WA_HEADER_UI-ZZ_DATBI","",ConvertDate(DT_WAK2_0600_BONUS_EXPIRATION_OCC2),False)
Call PressEnter()
'Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC3)
'Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC4)

Call SetTextbox("Bonus Expiration","RDM_S_BBY_WA_HEADER_UI-ZZ_DATBI","","",False)
Call PressEnter()
'Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC5)
'Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC6)


Call SetTextbox("Bonus Expiration","RDM_S_BBY_WA_HEADER_UI-ZZ_DATBI","",ConvertDate(DT_WAK2_0600_BONUS_EXPIRATION_OCC4),False)
Call PressEnter()
Call TakeScreenShot()
Call  ClickButton("Check Bonus Buy   \(Shift\+F7\)",False)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC8)

Call  ClickButton("Activate Bonus Buy   \(F8\)",False)
Call  ClickButton("Back   \(F3\)",False)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call ClickButton("Save   \(Ctrl\+S\)",False) 
Call VerifyStatusBarMessageType("S")
Call GetStatusBar("item1","DT_PROMOTION_NUMBER_OUTPUT")
Call VerifyStatusBar("Promotion "&DT_PROMOTION_NUMBER_OUTPUT&" saved")

Call LogOff()
Call FinalStatus ()








'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [12,8640008]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//




