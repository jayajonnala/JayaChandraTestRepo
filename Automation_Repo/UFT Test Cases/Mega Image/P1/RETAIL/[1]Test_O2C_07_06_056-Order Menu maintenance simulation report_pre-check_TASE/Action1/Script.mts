
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_O2C_07_06_056-Order Menu maintenance simulation report_pre-check_TASE
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

gstrTestCaseName = "Test_O2C_07_06_056-Order Menu maintenance simulation report_pre-check_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_O2C_07_06_056-Order Menu maintenance  simulation report_pre-check.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
Call SetTcode(DT_SAPTRANSACTIONCODE)     

Call PressEnter()

Call SetTextbox("Article","RMMW1-MATNR","",DT_MM43_0100_ARTICLE,False)
Call SetTextbox("Sales Org.","RMMW1-VKORG","",DT_MM43_0100_SALES_ORG,False)
Call SetTextbox("Distr. Channel","RMMW1-VTWEG","",DT_MM43_0100_DISTR_CHANNEL,False)
Call ClickButton("Deselect All   \(Shift\+F7\)",False)
Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100", "Screen description", "Sales", False)
Call PressEnter()
Call TakeScreenShot
Call GetTextboxValue("MVKE-MVGR4", "", "DT_MM43_2154_CHECK_TEXT_OF_ITEM_SET_OUTPUT", False)
Call VerifyTextBoxContent("Item Set", "MVKE-MVGR4", "", DT_MM43_2154_CHECK_TEXT_OF_ITEM_SET_OUTPUT, False)



''''''--------TransactionCode-/nZMDSO_ORDER_MENU----------''''

Call SetTcode(DT_MM43_4030_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Sales Organization","P_VKORG","",DT_MM43_1000_SALES_ORGANIZATION,False)
Call SetTextbox("Supplying Site","S_LOCLB-LOW","",DT_MM43_1000_SUPPLYING_SITE,False)
Call TakeScreenShot
Call SelectTab("TABSTRIP_GS_TAB", "Manual entry", False)
Call TakeScreenShot
Call SetTextbox("Customer no. - site","S_KUNNR-LOW","",DT_MM43_0200_CUSTOMER_NO__SITE,False)
Call SetTextbox("Distribution group","S_MVGR4-LOW","",DT_MM43_2154_CHECK_TEXT_OF_ITEM_SET_OUTPUT,False)
Call SetTextbox("Valid From/on","S_VALID-LOW","",ConvertDate(DT_MM43_0200_VALID_FROMON),False)
Call SetTextbox("to","S_VALID-HIGH","",ConvertDate(DT_MM43_0200_TO),False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot
Call SelectColumnGuiGrid("", 0, "DATBI", False)
Call ClickButtonToolBar("&MB_FILTER", 0)
Call SetTextbox("Valid To","%%DYN001-LOW","",ConvertDate(DT_MM43_1105_VALID_TO),True)
Call PressEnter()     
Call TakeScreenShot
Call GetGridContent("", 0, "ODSWD1", 1, "<NA>", "<NA>", "DT_MM43_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ODSWD1_OUTPUT")
Call GetGridContent("", 0, "ODSWD2", 1, "<NA>", "<NA>", "DT_MM43_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ODSWD2_OUTPUT")
Call GetGridContent("", 0, "ODSWD3", 1, "<NA>", "<NA>", "DT_MM43_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ODSWD3_OUTPUT")
Call GetGridContent("", 0, "ODSWD4", 1, "<NA>", "<NA>", "DT_MM43_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ODSWD4_OUTPUT")
Call GetGridContent("", 0, "ODSWD5", 1, "<NA>", "<NA>", "DT_MM43_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ODSWD5_OUTPUT")
Call GetGridContent("", 0, "ODSWD6", 1, "<NA>", "<NA>", "DT_MM43_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ODSWD6_OUTPUT")
Call GetGridContent("", 0, "ODSWD7", 1, "<NA>", "<NA>", "DT_MM43_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ODSWD7_OUTPUT")



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




