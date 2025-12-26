
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
'.................Test Script Name :Test_S2A_PRO_01_015-Create low level promotion New Store opening new tariff_P2
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_S2A_PRO_01_015- new tariff_P2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\smasu\Documents\TASE\Data Input\MI\DT_S2A_PRO_01_015-Create low level promotion New Store opening new tariff_P2.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)

Call PressEnter()

Call SetTcode(DT_SAPTRANSACTIONCODE)     

Call PressEnter()


Call SetTextBox("Promotion","WAKHD-AKTNR",0,DT_WAK5_1300_PROMOTION,False)

Call PressEnter()

Call TakeScreenShot()

Call ClickButton("MARKALL",False)

Call ClickButton("POSDETAIL",False)

Call CLickButton("Price activation   \(F6\)",True)

Call TakeScreenShot()

Call SelectMenuBar("Promotion;Check")

Call ClickLabel("4",0, False)

Call VerifyifGuiLabelExists(DT_WAK5_0120_CHECK_TEXT_OF_MT01)

Call ClickLabel("4",1,False)

Call VerifyifGuiLabelExists(DT_WAK5_0120_CHECK_TEXT_OF_MI_01_ARTEMIS)
Call VerifyifGuiLabelExists(DT_WAK5_0120_CHECK_TEXT_OF_MI_23_ARTEMIS)
Call VerifyifGuiLabelExists(DT_WAK5_0120_CHECK_TEXT_OF_MI_25_ARTEMIS)
Call VerifyifGuiLabelExists(DT_WAK5_0120_CHECK_TEXT_OF_MI_49_ARTEMIS)

Call ClickLabel("4",2, False)

Call VerifyifGuiLabelExists(DT_WAK5_0120_CHECK_TEXT_OF_7397874___________KG)
Call VerifyifGuiLabelExists(DT_WAK5_0120_CHECK_TEXT_OF_7410632___________EA)
Call VerifyifGuiLabelExists(DT_WAK5_0120_CHECK_TEXT_OF_7402024___________KG)
Call VerifyifGuiLabelExists(DT_WAK5_0120_CHECK_TEXT_OF_7377727___________KG)

Call ClickButton("Back   \(F3\)",False)

Call ClickButton("Save   \(Ctrl\+S\)",False)

Call TakeScreenShot()

Call GetStatusBar("item1","DT_DOCUMENT_NUMBER_OUTPUT")

Call VerifyStatusBar("Promotion "&DT_DOCUMENT_NUMBER_OUTPUT&" saved")


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




