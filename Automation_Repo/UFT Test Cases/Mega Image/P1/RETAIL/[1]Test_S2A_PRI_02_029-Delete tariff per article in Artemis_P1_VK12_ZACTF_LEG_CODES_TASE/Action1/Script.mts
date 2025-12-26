
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRI_02_029-Delete tariff per article in Artemis_P1_VK12_ZACTF_LEG_CODES 
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


gstrTestCaseName = "Test_S2A_PRI_02_029-DeleteP1_VK12_ZACTF_LEG_CODES"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_S2A_PRI_02_029-Delete tariff per article in Artemis_P2_ARTEMIS.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''''''''--------TransactionCode-VK12----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Condition Type","RV13A-KSCHL","",DT_VK12_0100_CONDITION_TYPE,False)
Call PressEnter()     
Call TakeScreenShot
Call SelectRadioButton("RV130-SELKZ", "Sales Org./Dist. Channel/Price List/Article/Sales Unit", True)
Call TakeScreenShot
Call ClickButtonIfExist("Choose   \(Enter\)",True)
Call TakeScreenShot
Call SetTextbox("Article","F004-LOW","",DT_VK12_1000_ARTICLE,False)
Call SetTextbox("Sales Organization","F001","",DT_VK12_1000_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution Channel","F002","",DT_VK12_1000_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Price List","F003","",DT_VK12_1000_PRICE_LIST,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call SelectRowGuiTable("SAPMV13ATCTRL_FAST_ENTRY","Article",DT_VK12_1000_ARTICLE,False)
'SAPGuiSession("Session").SAPGuiWindow("Change Sales price Condition").InsightObject("InsightObject").Click
' Updated Delete Row insight object
'SAPGuiSession("Session").SAPGuiWindow("DeleteRow").InsightObject("InsightObject").Click
Call ClickButton("Delete Row",False)
Call VerifyStatusBar(DT_VK12_1000_CHECK_TEXT_OF_STATUSBAR_DELETION)
Call VerifyStatusBarMessageType("S")
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call VerifyStatusBar(DT_VK12_1000_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot


''''''''--------TransactionCode-/nZACTF_LEG_CODES----------''''

Call SetTcode(DT_VK12_1000_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Group Name","P_GROUP","",DT_VK12_1000_GROUP_NAME,False)
Call SetTextbox("Article","S_MATNR-LOW","",DT_VK12_1000_ARTICLE_OCC1,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call GetGridContent("", 0, "Legacy Value", 1, "<NA>", "<NA>", "DT_VK12_0500_CHECK_GET_GRIDCELL_0_LEGACY_VALUE_OUTPUT")
Call TakeScreenShot

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


