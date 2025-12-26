
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRI_01_025-Enter additional import cost ZH22 to purchase price MEK1
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
'.................Test Script Name :Test_S2A_PRI_01_005-Enter purchase prices import article ZCCL - cost of labe
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_S2A_PRI_01_025- ZH22 to purchase price MEK1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_S2A_PRI_01_025-Enter additional import cost ZH22 to purchase pri.xls"
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
Call SetTextbox("Purch. Organization","KOMG-EKORG","",DT_MEK1_1893_PURCH_ORGANIZATION,False)
Call SetTextbox("Article","KOMG-MATNR","",DT_MEK1_1893_ARTICLE,False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Import", 1, "", "", DT_MEK1_1893_TABLECELL_IMPORT_0, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Amount", 1, "", "", DT_MEK1_1893_TABLECELL_AMOUNT_0, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Unit", 1, "", "", DT_MEK1_1893_TABLECELL_UNIT_0, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "UoM", 1, "", "", DT_MEK1_1893_TABLECELL_UOM, False)
Call PressEnter()     
Call TakeScreenShot
Call ClickButton("Select All",False)
Call SelectMenuBar("Goto;Details")
''Call SetTextbox("Vendor","KONP-LIFNR","",DT_MEK1_0300_VENDOR,False)
Call SetTextboxNoLabel("KONP-LIFNR","",DT_MEK1_0300_VENDOR,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call TakeScreenShot
Call VerifyStatusBar(DT_MEK1_1893_CHECK_TEXT_OF_STATUSBAR)
Call SelectMenuBar("Condition;Display")
Call SetTextbox("Purch. Organization","F001","",DT_MEK1_1000_PURCH_ORGANIZATION,False)
Call SetTextbox("Article","F002","",DT_MEK1_1000_ARTICLE,False)
Call SetTextbox("Valid On","SEL_DATE","",ConvertDate(DT_MEK1_1000_VALID_ON),False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call VerifyTextBoxContent("Article","KOMG-MATNR","",DT_MEK1_1893_CHECK_TEXT_OF_ARTICLE,False)
Call VerifyTableCellContent(1, "Import", "SAPMV13ATCTRL_FAST_ENTRY", DT_MEK1_1893_CHECK_TEXT_OF_TABLECELL_IMPORT_0)
Call VerifyTableCellContent(1, "Amount", "SAPMV13ATCTRL_FAST_ENTRY", DT_MEK1_1893_CHECK_TEXT_OF_TABLECELL_AMOUNT_0)
Call VerifyTableCellContent(1, "Unit", "SAPMV13ATCTRL_FAST_ENTRY", DT_MEK1_1893_CHECK_TEXT_OF_TABLECELL_UNIT_0)


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




