
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_O2C_08_07_098 -Faire share-substitution_P1_Substitute_Article     
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


gstrTestCaseName = "Test_O2C_08_07_098 _P1_Substitute_Article"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_O2C_08_07_098 -Faire share-substitution_P1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


''''''--------TransactionCode-WRFFUART----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call ClickButton("Article Selection   \(F6\)", False)
Call SetTextbox("Article","S_ARTNR-LOW", "",DT_WRFFUART_0103_ARTICLE, True)
Call SetTextbox("Substit. Profile","P_FUTYP1", "", DT_WRFFUART_0103_SUBSTIT_PROFILE, True)
Call TakeScreenShot
Call PressEnter()
Call ClickButton("Execute   \(F8\)", False)
Call SelectAllRowGuiGrid("Article Selection", 0, False)
Call ClickButtonToolBar("ASSIGN", 0)
Call VerifyGridCellContent("Substitution Type: 90 Stock Driven Substitution", 1, "FOLUP_ART_NR", "", DT_WRFFUART_0100_GRIDCELL_0_FUREPLACEMENT_ART)
Call SelectAllRowGuiGrid("Substitution Type: 90 Stock Driven Substitution",0, False)
''Call ClickButtonToolBar("DELE", 0)

Call Click204ButtonToolBar("DELE", 1)


Call ClickButton("btn\[11\]", False)  ''not required
Call ClickButtonIfExist("Yes", True)
Call ClickButtonIfExist("btn\[0\]", True)
Call ClickButtonIfExist("btn\[0\]", True)

Call ClickButtonToolBar("ASSIGN", "") @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf1.xml_;_
Call SetGridData("Substitution Type: 90 Stock Driven Substitution", 1, "ORG_ART_FACTOR", DT_WRFFUART_0100_GRIDCELL_0_FACTOR, False)
Call SetGridData("Substitution Type: 90 Stock Driven Substitution", 1, "SUBST_ART_FACTOR", DT_WRFFUART_0100_GRIDCELL_0_FACTOR_OCC1, False)
Call SetGridData("Substitution Type: 90 Stock Driven Substitution", 1, "FOLUP_ART_NR", DT_WRFFUART_0100_GRIDCELL_0_FUREPLACEMENT_ART, False)
Call SetGridData("Substitution Type: 90 Stock Driven Substitution", 1, "ASORT", DT_WRFFUART_0100_GRIDCELL_0_ASSORTMENT, False)
Call PressEnter()
Call TakeScreenShot
Call SelectCheckBoxGridByRowNo("Substitution Type: 90 Stock Driven Substitution","Follow-Up Action", 1, "", "ON")

Call ClickButton("Save   \(Ctrl\+S\)", False)

''Call VerifyTextPopup("Follow-up/replacement article relationships posted")
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)", True)

 @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf3.xml_;_

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
 @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf53.xml_;_
 @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf53.xml_;_
