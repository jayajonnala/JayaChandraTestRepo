
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRI_02_023-Maintain margin driven retail price for tariff 49 ShopandGo_SAP 
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

gstrTestCaseName = "Test_S2A_PRI_02_023- 49 ShopandGo_SAP"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_S2A_PRI_02_023-Maintain margin driven retail price for tariff 49 ShopandGo_SAP.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

''''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''''''''--------TransactionCode-VKP5----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call ClickButton("Get Variant...   \(Shift\+F5\)",False)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call ClickButton("Execute   \(F8\)",True)
Call ClickButtonToolBar("&FIND", 0)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_VKP5_0841_SEARCH_TERM,True)
Call TakeScreenShot
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
''grid title has been changed
''Call SelectRowGuiGrid("Variant Catalog for Program RWVKP007", 0, "Variant name", DT_VKP5_0841_SEARCH_TERM, True)
Call SelectRowGuiGrid("Variant Catalog.*", 0, "Variant name", DT_VKP5_0841_SEARCH_TERM, True)
Call ClickButton("Choose   \(F2\)",True)
Call SetTextbox("Article","S_MATNR-LOW","","",False)
Call SetTextbox("Validity","S_DATUM-LOW","",ConvertDate(DT_VKP5_1000_VALIDITY),False)
Call SetTextbox("to","S_DATUM-HIGH","",ConvertDate(DT_VKP5_1000_TO),False)
Call TakeScreenShot
Call ClickButton("%_S_MATNR_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "<NA>", "<NA>", DT_VKP5_3010_TABLECELL_SINGLE_VALUE_0, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "<NA>", "<NA>", DT_VKP5_3010_TABLECELL_SINGLE_VALUE_1, True)
Call ClickButton("Copy   \(F8\)",True)
Call TakeScreenShot
Call SetTextbox("Sales organization","S_VKORG-LOW","",DT_VKP5_1000_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution channel","S_VTWEG-LOW","",DT_VKP5_1000_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Price List","S_PLTYP-LOW","",DT_VKP5_1000_PRICE_LIST,False)
Call TakeScreenShot
Call SetTextbox("Purchase Price Determ. Seq.","P_EKERV","",DT_PRICE_PURCH_DETERM,False)
Call SetTextbox("Sales price determination seq.","P_VKERV","",DT_SALES_PRICE_DETERM,False)
Call SetTextbox("List Group","P_LIGRU","",DT_LIST_GROUP,False)
Call SetTextbox("List Variant","P_LIVAR","",DT_LIST_VARIANT,False)
Call PressEnter()     
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call SelectColumnGuiGrid("", 0, "PLTYP", False)
Call ClickButtonToolBar("&MB_FILTER", 0)
Call SetTextbox("Price List","%%DYN001-LOW","",DT_VKP5_1000_PRICE_LIST,True)
Call TakeScreenShot
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot
Call GetGridContent("", 0, "ENDPA", 1, "<NA>", "<NA>", "DT_VKP5_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ENDPA_OUTPUT")
Call GetGridContent("", 0, "ENDPA", 2, "<NA>", "<NA>", "DT_VKP5_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ENDPA_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetGridData("", 1, "ENDPR", DT_VKP5_0100_GRIDCELL_0_FINAL_PRICE, False)
Call SetGridData("", 2, "ENDPR",DT_VKP5_0100_GRIDCELL_1_FINAL_PRICE, False)

''Call SetGridData("", 1, "ENDPR", CDec(DT_VKP5_0100_GRIDCELL_0_FINAL_PRICE_OCC1)+0.06, False)
''Call SetGridData("", 2, "ENDPR", Cint(DT_VKP5_0100_GRIDCELL_1_FINAL_PRICE_OCC1)+0.06, False)
Call WriteRunTimeDataToExcelGlobalSheet(DT_VKP5_0100_GRIDCELL_0_FINAL_PRICE_OCC2, Cstr(Cint(DT_VKP5_0100_GRIDCELL_0_FINAL_PRICE_OCC1)+0.06))
Call WriteRunTimeDataToExcelGlobalSheet(DT_VKP5_0100_GRIDCELL_1_FINAL_PRICE_OCC2, Cstr(Cint(DT_VKP5_0100_GRIDCELL_1_FINAL_PRICE_OCC1)+0.06))

Call PressEnter()     
Call TakeScreenShot
Call SelectAllRowGuiGrid("", 0, False)
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call TakeScreenShot
Call GetStatusBar("item1", "DT_VKP5_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Data saved; pricing document "&DT_VKP5_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" created")
Call ClickButtonIfExist("Continue   \(Enter\)",True)
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


