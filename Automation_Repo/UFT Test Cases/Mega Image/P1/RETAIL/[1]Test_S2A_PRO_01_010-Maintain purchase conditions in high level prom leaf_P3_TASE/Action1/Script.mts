
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRO_01_010-Maintain purchase conditions in high level prom leaf_P3
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
'.................Test Script Name :Test_S2A_PRO_01_010-Maintain purchase conditions in high level prom leaf_P3
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


gstrTestCaseName = "Test_S2A_PRO_01_010-prom leaf_P3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_S2A_PRO_01_010-Maintain purchase conditions in high level prom leaf_P3.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)

Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()



''''''''''''''''''''''' TCODE MIGO '''''''''''''''''''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter() 
Call TakeScreenShot()
'
'Call SetTextBox("Condition Type","RV13A-KSCHL",0,DT_MEK3_0100_CONDITION_TYPE,False)
'Call ClickButton("Condition Information   \(Shift\+F4\)",False)
'Call ClickButton("%_F001_%_APP_%-VALU_PUSH",False)
'Call SetTableData("SAPLALDBSINGLE","Single Value", 1, "", "", DT_MEK3_3010_TABLECELL_SINGLE_VALUE_0, True)
'Call SetTableData("SAPLALDBSINGLE","Single Value", 2, "", "", DT_MEK3_3010_TABLECELL_SINGLE_VALUE_1, True)
'Call SetTableData("SAPLALDBSINGLE","Single Value", 3, "", "", DT_MEK3_3010_TABLECELL_SINGLE_VALUE_2, True)
'Call SetTableData("SAPLALDBSINGLE","Single Value", 4, "", "", DT_MEK3_3010_TABLECELL_SINGLE_VALUE_3, True)
'Call SetTableData("SAPLALDBSINGLE","Single Value", 5, "", "", DT_MEK3_3010_TABLECELL_SINGLE_VALUE_4, True)
'Call SetTableData("SAPLALDBSINGLE","Single Value", 6, "", "", DT_MEK3_3010_TABLECELL_SINGLE_VALUE_5, True)
'Call SetTableData("SAPLALDBSINGLE","Single Value", 7, "", "", DT_MEK3_3010_TABLECELL_SINGLE_VALUE_6, True)
'Call SetTableData("SAPLALDBSINGLE","Single Value", 8, "", "", DT_MEK3_3010_TABLECELL_SINGLE_VALUE_7, True)
'Call ClickButton("Copy   \(F8\)",True)
'Call ClickButton("%_F002_%_APP_%-VALU_PUSH",False)
'Call ClickButton("Delete Entire Selection   \(Shift\+F4\)",True)
'Call SetTableData("SAPLALDBSINGLE","Single Value", 1, "", "", DT_MEK3_3010_TABLECELL_SINGLE_VALUE_0_OCC1, True)
'Call SetTableData("SAPLALDBSINGLE","Single Value", 2, "", "", DT_MEK3_3010_TABLECELL_SINGLE_VALUE_1_OCC1, True)
'Call SetTableData("SAPLALDBSINGLE","Single Value", 3, "", "", DT_MEK3_3010_TABLECELL_SINGLE_VALUE_2_OCC1, True)
'Call SetTableData("SAPLALDBSINGLE","Single Value", 4, "", "", DT_MEK3_3010_TABLECELL_SINGLE_VALUE_3_OCC1, True)
'Call SetTableData("SAPLALDBSINGLE","Single Value", 5, "", "", DT_MEK3_3010_TABLECELL_SINGLE_VALUE_4_OCC1, True)
'Call SetTableData("SAPLALDBSINGLE","Single Value", 6, "", "", DT_MEK3_3010_TABLECELL_SINGLE_VALUE_5_OCC1, True)
'Call SetTableData("SAPLALDBSINGLE","Single Value", 7, "", "", DT_MEK3_3010_TABLECELL_SINGLE_VALUE_6_OCC1, True)
'Call SetTableData("SAPLALDBSINGLE","Single Value", 8, "", "", DT_MEK3_3010_TABLECELL_SINGLE_VALUE_7_OCC1, True)
'Call ClickButton("Copy   \(F8\)",True)
'
'Call SetTextBox("Purch\. Organization","F003-LOW",0,DT_MEK3_1000_PURCH_ORGANIZATION,False)
'Call SetTextBox("Info record category","F005-LOW",0,DT_MEK3_1000_INFO_RECORD_CATEGORY,False)
'Call SetTextBox("from / on","SEL_DATE",0,ConvertDate(DT_MEK3_1000_FROM__ON),False)
'Call ClickButton("btn\[8\]",False)
'Call TakeScreenShot()
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_RP01,0)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_RP01_OCC1,1)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_RP01_OCC2,2)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_RP01_OCC3,3)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_RP01_OCC4,4)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_RP01_OCC5,5)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_RP01_OCC6,6)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_RP01_OCC7,7)
'
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME,0)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC1,1)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC2,0)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC3,1)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC4,2)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC5,3)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC6,4)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC7,0)
'
'
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC8,0)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC9,1)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC10,2)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC11,3)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC12,4)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC13,5)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC14,6)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC15,7)
'
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC16,0)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC17,1)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC18,2)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC19,3)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC20,4)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC21,5)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC22,6)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC23,7)
'
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC24,0)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC25,1)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC26,2)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC27,3)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC28,4)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC29,5)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC30,6)
'Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC31,7)

'Call ClickButton("Back   \(F3\)",False)
'Call ClickButton("Back   \(F3\)",False)

Call SetTextBox("Condition Type","RV13A-KSCHL",0,DT_MEK3_0100_CONDITION_TYPE_OCC1,False)
Call ClickButton("Condition Information   \(Shift\+F4\)",False)
Call ClickButton("%_F001_%_APP_%-VALU_PUSH",False)
Call ClickButton("Delete Entire Selection   \(Shift\+F4\)",True)
Call SetTableData("SAPLALDBSINGLE","Single Value", 1, "", "", DT_MEK3_3010_TABLECELL_SINGLE_VALUE_0_OCC2, True)
Call ClickButton("Copy   \(F8\)",True)
Call ClickButton("%_F002_%_APP_%-VALU_PUSH",False)
Call ClickButton("Delete Entire Selection   \(Shift\+F4\)",True)
Call SetTableData("SAPLALDBSINGLE","Single Value", 1, "", "", DT_MEK3_3010_TABLECELL_SINGLE_VALUE_0_OCC3, True)
Call ClickButton("Copy   \(F8\)",True)

Call SetTextBox("Purch\. Organization","F003-LOW",0,DT_MEK3_1000_PURCH_ORGANIZATION_OCC1,False)
Call SetTextBox("Info record category","F005-LOW",0,DT_MEK3_1000_INFO_RECORD_CATEGORY_OCC1,False)
Call SetTextBox("from / on","SEL_DATE",0,ConvertDate(DT_MEK3_1000_FROM__ON_OCC1),False)
Call ClickButton("btn\[8\]",False)
Call TakeScreenShot()
Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_RP01_OCC8,0)
Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC32,0)
Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC33,0)
Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC34,0)
Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC35,0)


Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)

Call SetTextBox("Condition Type","RV13A-KSCHL",0,DT_MEK3_0100_CONDITION_TYPE_OCC2,False)
Call ClickButton("Condition Information   \(Shift\+F4\)",False)
Call ClickButton("%_F001_%_APP_%-VALU_PUSH",False)
Call ClickButton("Delete Entire Selection   \(Shift\+F4\)",True)
Call SetTableData("SAPLALDBSINGLE","Single Value", 1, "", "", DT_MEK3_3010_TABLECELL_SINGLE_VALUE_0_OCC4, True)
Call SetTableData("SAPLALDBSINGLE","Single Value", 2, "", "", DT_MEK3_3010_TABLECELL_SINGLE_VALUE_1_OCC2, True)
Call ClickButton("Copy   \(F8\)",True)
Call ClickButton("%_F002_%_APP_%-VALU_PUSH",False)
Call ClickButton("Delete Entire Selection   \(Shift\+F4\)",True)
Call SetTableData("SAPLALDBSINGLE","Single Value", 1, "", "", DT_MEK3_3010_TABLECELL_SINGLE_VALUE_0_OCC5, True)
Call SetTableData("SAPLALDBSINGLE","Single Value", 2, "", "", DT_MEK3_3010_TABLECELL_SINGLE_VALUE_1_OCC3, True)
Call ClickButton("Copy   \(F8\)",True)

Call SetTextBox("Purch\. Organization","F003-LOW",0,DT_MEK3_1000_PURCH_ORGANIZATION_OCC2,False)
Call SetTextBox("Info record category","F005-LOW",0,DT_MEK3_1000_INFO_RECORD_CATEGORY_OCC2,False)
Call SetTextBox("from / on","SEL_DATE",0,ConvertDate(DT_MEK3_1000_FROM__ON_OCC2),False)
Call ClickButton("btn\[8\]",False)
Call TakeScreenShot()
Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_RP01_OCC9,0)
Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_RP01_OCC10,1)
Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC36,0)
Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC37,1)
Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC38,0)
Call VerifyIfGuilabelExists_ByIndex(DT_MEK3_0120_CHECK_TEXT_OF_NO_NAME_OCC39,1)


Call LogOff()
Call FinalStatus()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''










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




