


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
'.................Test Script Name :Test_01PRO04_002_Add_Purch.Condition_ZPD1_(cumulat)
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_01PRO04_002_Add_Purch.Condition_ZPD1_(cumulat)"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRO04_002_Add_Purch.Condition_ZPD1_(cumulat)_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =4
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

'--------------------------------------------  WAK3------------------------------------------------

Call SetTextbox("Promotion","WAKHD-AKTNR","",DT_WAK2_1100_PROMOTION,False)
Call TakeScreenShot()
Call  ClickButton("Fast entry   \(F5\)",False)
Call TakeScreenShot()
Call SelectTab("TAXI_TABSTRIP_UEBERSICHT","Pricing conditions (purchasing)",False)
Call TakeScreenShot()
Call  ClickButton("Create condition",False)
Call TakeScreenShot()


Call SetVerticalScrollBar(45,True)  '---- Scroll down is used to make Article Info Record \(NEW\) label visible while execution
Call ClickLabel("Article Info Record \(NEW\)",0,True)

Call ClickButtonIfExist("Choose   \(F2\)",True)

Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Supplier",1,"","",DT_WAK2_1018_VENDOR,False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Article",1,"","",DT_WAK2_1018_ARTICLE,False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Amount",1,"","",DT_WAK2_1018_TABLECELL_AMOUNT_0,False)
Call SetTextbox("Purch\. Organization","KOMG-EKORG","",DT_WAK2_1018_PURCH_ORGANIZATION,False)
Call SetTextbox("Info record category","KOMG-ESOKZ","",DT_WAK2_1018_TABLECELL_INFOTYPE_0,False)
Call PressEnter()
Call TakeScreenShot()
Call ClickButton("Save   \(Ctrl\+S\)",False) 
Call VerifyStatusBarMessageType("S")
Call VerifyStatusBar("Promotion "&DT_WAK2_1100_PROMOTION&" saved")


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




