
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRI_01_062-Set purchase condittion ZREB
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
'.................Test Script Name :Test_S2A_PRI_01_062-Set purchase condittion ZREB
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_S2A_PRI_01_062-Set purchase condittion ZREB"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_S2A_PRI_01_062-Set purchase condittion ZREB.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)

Call Login(DT_SAPUSER,DT_SAPPASSWORD)

Call PressEnter()

Call SetTcode(DT_SAPTRANSACTIONCODE)     

Call PressEnter()   
Call SetTextbox("Condition Type","RV13A-KSCHL","",DT_MEK1_0100_CONDITION_TYPE,False)
Call ClickButton("Key Combination   \(Shift\+F5\)",False)
Call ClickButton("Choose   \(Enter\)",True)
Call SetTextbox("Purch. Organization","KOMG-EKORG","",DT_MEK1_1992_PURCH_ORGANIZATION,False)
Call SetTextboxNoLabel("KOMG-LIFNR","",DT_MEK1_1992_VENDOR,False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Article", 1, "", "", DT_MEK1_1992_TABLECELL_ARTICLE_0, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Amount", 1, "", "", DT_MEK1_1992_TABLECELL_AMOUNT_0, False)
Call PressEnter()   
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call TakeScreenShot
Call VerifyStatusBar(DT_MEK1_1992_CHECK_TEXT_OF_STATUSBAR)

'''--------TransactionCode-MEK3----------''''

Call SetTcode(DT_MEK1_1992_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Condition Type","RV13A-KSCHL","",DT_MEK1_0100_CONDITION_TYPE_OCC1,False)
Call ClickButton("Condition Information   \(Shift\+F4\)",False)
Call SetTextbox("Purch. Organization","F001-LOW","",DT_MEK1_1000_PURCH_ORGANIZATION,False)
Call SetTextboxNoLabel("F002-LOW","",DT_MEK1_1000_VENDOR,False)
Call SetTextbox("Article","F003-LOW","",DT_MEK1_1000_ARTICLE,False)
Call SetTextbox("from / on","SEL_DATE","",ConvertDate(DT_MEK1_1000_FROM__ON),False)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call ClickButton("Select All   \(F7\)",False)
Call ClickButton("Display   \(F5\)",False)
Call TakeScreenShot
'''Call VerifyTextBoxContent("Vendor","KOMG-LIFNR","",DT_MEK1_1992_CHECK_TEXT_OF_VENDOR,False)

Call VerifyTextBoxNoLabelContent("KOMG-LIFNR","",DT_MEK1_1992_CHECK_TEXT_OF_VENDOR,False)

Call VerifyTableCellContent(1, "Article", "SAPMV13ATCTRL_FAST_ENTRY", DT_MEK1_1992_CHECK_TEXT_OF_TABLECELL_ARTICLE_0)
Call VerifyTableCellContent(1, "Amount", "SAPMV13ATCTRL_FAST_ENTRY", DT_MEK1_1992_CHECK_TEXT_OF_TABLECELL_AMOUNT_0)



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




