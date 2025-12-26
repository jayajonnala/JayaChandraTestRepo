
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRO_01_010-Maintain purchase conditions in high level prom leaf_P2
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
'.................Test Script Name :Test_P2P_01_01_0113-Number Range for manual GR done in SAP_P1_GR
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


gstrTestCaseName = "Test_S2A_PRO_01_010-prom leaf_P2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_S2A_PRO_01_010-Maintain purchase conditions in high level prom leaf_P2.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)

Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'''''''''''''''''''''''' TCODE MIGO '''''''''''''''''''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter() 
Call TakeScreenShot()
Call SetTextBox("Promotion","WAKHD-AKTNR",0,DT_WAK2_1100_PROMOTION,False)
Call PressEnter()
Call TakeScreenShot()
Call SelectTab("TAXI_TABSTRIP_UEBERSICHT","Pricing conditions (purchasing)",False)
Call GetTextboxValue("WAKHD-EKDAB", 0, "DT_PURCHASE_FROM_OUTPUT", False)
Call GetTextboxValue("WAKHD-EKDBI",0,"DT_PURCHASE_TO_OUTPUT",False)
Call ClickButton("EKKONDANL",False)
Call SetFocusGuiLabel(DT_WAK2_0800_FIND,11,56,True)
Call ClickButton("Choose   \(F2\)",True)
Call TakeScreenShot()
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Supplier", 1, "", "", DT_WAK2_1906_TABLECELL_VENDOR_0, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Supplier", 2, "", "", DT_WAK2_1906_TABLECELL_VENDOR_1, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Supplier", 3, "", "", DT_WAK2_1906_TABLECELL_VENDOR_2, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Supplier", 4, "", "", DT_WAK2_1906_TABLECELL_VENDOR_3, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Supplier", 5, "", "", DT_WAK2_1906_TABLECELL_VENDOR_4, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Supplier", 6, "", "", DT_WAK2_1906_TABLECELL_VENDOR_5, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Supplier", 7, "", "", DT_WAK2_1906_TABLECELL_VENDOR_6, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Supplier", 8, "", "", DT_WAK2_1906_TABLECELL_VENDOR_7, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Article", 1, "", "", DT_WAK2_1906_TABLECELL_ARTICLE_0, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Article", 2, "", "", DT_WAK2_1906_TABLECELL_ARTICLE_1, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Article", 3, "", "", DT_WAK2_1906_TABLECELL_ARTICLE_2, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Article", 4, "", "", DT_WAK2_1906_TABLECELL_ARTICLE_3, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Article", 5, "", "", DT_WAK2_1906_TABLECELL_ARTICLE_4, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Article", 6, "", "", DT_WAK2_1906_TABLECELL_ARTICLE_5, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Article", 7, "", "", DT_WAK2_1906_TABLECELL_ARTICLE_6, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Article", 8, "", "", DT_WAK2_1906_TABLECELL_ARTICLE_7, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Amount", 1, "", "", DT_WAK2_1906_TABLECELL_AMOUNT_0, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Amount", 2, "", "", DT_WAK2_1906_TABLECELL_AMOUNT_1, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Amount", 3, "", "", DT_WAK2_1906_TABLECELL_AMOUNT_2, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Amount", 4, "", "", DT_WAK2_1906_TABLECELL_AMOUNT_3, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Amount", 5, "", "", DT_WAK2_1906_TABLECELL_AMOUNT_4, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Amount", 6, "", "", DT_WAK2_1906_TABLECELL_AMOUNT_5, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Amount", 7, "", "", DT_WAK2_1906_TABLECELL_AMOUNT_6, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Amount", 8, "", "", DT_WAK2_1906_TABLECELL_AMOUNT_7, False)
Call SetTextBox("Purch\. Organization","KOMG-EKORG",0,DT_WAK2_1906_PURCH_ORGANIZATION,False)
Call SetTextBox("Info record category","KOMG-ESOKZ",0,DT_WAK2_1906_INFO_RECORD_CATEGORY,False)
Call PressEnter()
Call TakeScreenShot()
Call ClickButton("Save   \(Ctrl\+S\)",False)

'Call PressEnter()
'Call PressEnter()
'Call PressEnter()
'Call PressEnter()
'Call PressEnter()
'Call PressEnter()
'Call PressEnter()
'Call PressEnter()

Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)

Call TakeScreenShot()
Call GetStatusBar("text","DT_WAK2_1100_GET_TEXT_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar(DT_WAK2_1100_GET_TEXT_OF_STATUSBAR_OUTPUT)
Call SetTextBox("Promotion","WAKHD-AKTNR",0,DT_WAK2_1100_PROMOTION_OCC1,False)
Call PressEnter()
Call TakeScreenShot()
Call SelectTab("TAXI_TABSTRIP_UEBERSICHT","Pricing conditions (purchasing)",False)

Call ClickButton("EKKONDANL",False)
Call SetFocusGuiLabel(DT_WAK2_0800_FIND_OCC1,11,56,True)
Call ClickButton("Choose   \(F2\)",True)
Call TakeScreenShot()
Call SetTextBox("Purch\. Organization","KOMG-EKORG",0,DT_WAK2_1906_PURCH_ORGANIZATION_OCC1,False)
Call SetTextBox("Info record category","KOMG-ESOKZ",0,DT_WAK2_1906_INFO_RECORD_CATEGORY_OCC1,False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Supplier", 1, "", "", DT_WAK2_1906_TABLECELL_VENDOR_0_OCC1, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Article", 1, "", "", DT_WAK2_1906_TABLECELL_ARTICLE_0_OCC1, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Amount", 1, "", "", DT_WAK2_1906_TABLECELL_AMOUNT_0_OCC1, False)
Call PressEnter()
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call GetStatusBar("text","DT_WAK2_1100_GET_TEXT_OF_STATUSBAR_OCC2_OUTPUT")
Call VerifyStatusBar(DT_WAK2_1100_GET_TEXT_OF_STATUSBAR_OCC2_OUTPUT)

Call SetTextBox("Promotion","WAKHD-AKTNR",0,DT_WAK2_1100_PROMOTION_OCC2,False)
Call PressEnter()
Call TakeScreenShot()
Call SelectTab("TAXI_TABSTRIP_UEBERSICHT","Pricing conditions (purchasing)",False)
Call ClickButton("EKKONDANL",False)
Call SetFocusGuiLabel(DT_WAK2_0800_FIND_OCC2,11,56,True)
Call ClickButton("Choose   \(F2\)",True)
Call TakeScreenShot()
Call SetTextBox("Purch\. Organization","KOMG-EKORG",0,DT_WAK2_1906_PURCH_ORGANIZATION_OCC2,False)
Call SetTextBox("Info record category","KOMG-ESOKZ",0,DT_WAK2_1906_INFO_RECORD_CATEGORY_OCC2,False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Supplier", 1, "", "", DT_WAK2_1906_TABLECELL_VENDOR_0_OCC2, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Supplier", 2, "", "", DT_WAK2_1906_TABLECELL_VENDOR_1_OCC1, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Article", 1, "", "", DT_WAK2_1906_TABLECELL_ARTICLE_0_OCC2, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Article", 2, "", "", DT_WAK2_1906_TABLECELL_ARTICLE_1_OCC1, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Amount", 1, "", "", DT_WAK2_1906_TABLECELL_AMOUNT_0_OCC2, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Amount", 2, "", "",DT_WAK2_1906_TABLECELL_AMOUNT_1_OCC1, False)
Call PressEnter()
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call GetStatusBar("text","DT_WAK2_1100_GET_TEXT_OF_STATUSBAR_OCC4_OUTPUT")
Call VerifyStatusBar(DT_WAK2_1100_GET_TEXT_OF_STATUSBAR_OCC4_OUTPUT)

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


 @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf53.xml_;_
