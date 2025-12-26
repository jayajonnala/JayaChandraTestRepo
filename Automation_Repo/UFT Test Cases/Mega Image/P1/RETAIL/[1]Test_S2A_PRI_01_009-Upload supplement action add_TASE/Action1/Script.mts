
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRI_01_009-Upload supplement action add 
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

gstrTestCaseName = "Test_S2A_PRI_01_009-Upload supplement action add"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_S2A_PRI_01_009-Upload supplement action add.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''
'
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'''--------TransactionCode-ZMDPC_UPLOAD_COND----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call ClickButton("Get Variant...   \(Shift\+F5\)",False)
Call SetTextbox("Variant","V-LOW","",DT_ZMDPC_UPLOAD_COND_0100_VARIANT,True)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call ClickButton("Execute   \(F8\)",True)
Call ClickButtonToolBar("&FIND", 0)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_ZMDPC_UPLOAD_COND_0841_SEARCH_TERM,True)
Call TakeScreenShot
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call SelectRowGuiGrid("Variant Catalog for Program ZMDPC_UPLOAD_PRICES_FROM_FILE"&".*", 0, "Variant name", DT_ZMDPC_UPLOAD_COND_0841_SEARCH_TERM, True)
Call ClickButton("Choose   \(F2\)",True)
Call SetTextbox("Supplement Condition Type","P_KSCHLS","",DT_ZMDPC_UPLOAD_COND_1000_SUPPLEMENT_CONDITION_TYPE,False)
Call SetComboByKey("Supplement action", DT_ZMDPC_UPLOAD_COND_1000_SUPPLEMENT_ACTION)
Call SetTextbox("File name","P_FILE","","",False)
Call TakeScreenShot
Call FocusTextBox("File name", "P_FILE", False)
Call SendKey("{F4}")
Wait(5)
Call SetTextbox("Directory","DY_PATH","",DT_ZMDPC_UPLOAD_COND_0200_DIRECTORY,True)
Call SetTextbox("File Name","DY_FILENAME","",DT_ZMDPC_UPLOAD_COND_0200_FILE_NAME,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call ClickButton("Select All   \(F5\)",False)
Call ClickButton("Create Conditions   \(F8\)",False)
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "Exception", 0, DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LIGHT)
Call VerifyGridCellContent("", 2, "Exception", 0, DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LIGHT)
Call VerifyGridCellContent("", 3, "Exception", 0, DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_LIGHT)
Call GetGridContent("", 0, "MATNR", 1, "<NA>", "<NA>", "DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR_OUTPUT")

''''--------TransactionCode-/nMM43----------''''

Call SetTcode(DT_ZMDPC_UPLOAD_COND_0500_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Article","RMMW1-MATNR","",DT_ZMDPC_UPLOAD_COND_0100_ARTICLE,false)
Call SetTextbox("Purchasing Org.","RMMW1-EKORG","",DT_ZMDPC_UPLOAD_COND_0100_PURCHASING_ORG,false)
call SetTextboxNoLabel("RMMW1-LIFNR",2,DT_ZMDPC_UPLOAD_COND_0100_VENDOR,false)
'Call SetTextbox("Vendor","RMMW1-LIFNR","",DT_ZMDPC_UPLOAD_COND_0100_VENDOR,false)  '' Sandbox changes from vendor to supplier
Call ClickButton("Deselect All   \(Shift\+F7\)",False)
Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100", "Screen description", "Purchasing", False)
Call PressEnter() 
Call ClickButton("For the Conditions",False)
Call TakeScreenShot
Call ClickButton("btn\[24\]",True)
Call SelectCellGuiTable("SAPLV14ATCTRL_D0102", "Valid From", "Valid to", "31.12.9999", True)
Call ClickButton("btn\[0\]",True)
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "Supplier", 0, DT_ZMDPC_UPLOAD_COND_0010_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LIFNR)
'Call VerifyGridCellContent("", 1, "Vendor", 0, DT_ZMDPC_UPLOAD_COND_0010_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LIFNR)  '' Sandbox changes from vendor to supplier
Call VerifyGridCellContent("", 1, "Article", 0, DT_ZMDPC_UPLOAD_COND_0010_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR)
Call GetTableCellData("SAPMV13ATCTRL_D0201", "Amount", 3, "<NA>", "<NA>", "DT_GET_OUTPUT", False)
Call VerifyTableCellContent(3, "Amount", "SAPMV13ATCTRL_D0201", DT_GET_OUTPUT)


''--------TransactionCode-ZMDPC_UPLOAD_COND----------''''

Call SetTcode(DT_ZMDPC_UPLOAD_COND_0201_OKCD)     
Call PressEnter()     

Call ClickButton("Get Variant...   \(Shift\+F5\)",False)
Call SetTextbox("Variant","V-LOW","",DT_ZMDPC_UPLOAD_COND_0100_VARIANT_OCC1,True)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call ClickButton("Execute   \(F8\)",True)
Call ClickButtonToolBar("&FIND", 0)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_ZMDPC_UPLOAD_COND_0841_SEARCH_TERM_OCC1,True)
Call TakeScreenShot
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call SelectRowGuiGrid("Variant Catalog for Program ZMDPC_UPLOAD_PRICES_FROM_FILE"&".*", 0, "Variant name", DT_ZMDPC_UPLOAD_COND_0841_SEARCH_TERM_OCC1, True)
Call ClickButton("Choose   \(F2\)",True)
Call SetTextbox("Supplement Condition Type","P_KSCHLS","",DT_ZMDPC_UPLOAD_COND_1000_SUPPLEMENT_CONDITION_TYPE_OCC1,False)
Call SetComboByKey("Supplement action", DT_ZMDPC_UPLOAD_COND_1000_SUPPLEMENT_ACTION_OCC1)
Call SetTextbox("File name","P_FILE","","",False)
Call TakeScreenShot
Call FocusTextBox("File name", "P_FILE", False)
Call SendKey("{F4}")
Wait(5)
Call SetTextbox("Directory","DY_PATH","",DT_ZMDPC_UPLOAD_COND_0200_DIRECTORY_OCC1,True)
Call SetTextbox("File Name","DY_FILENAME","",DT_ZMDPC_UPLOAD_COND_0200_FILE_NAME_OCC1,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call ClickButton("Select All   \(F5\)",False)
Call ClickButton("Create Conditions   \(F8\)",False)
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "Exception", 0, DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LIGHT)
Call VerifyGridCellContent("", 2, "Exception", 0, DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LIGHT)
Call VerifyGridCellContent("", 3, "Exception", 0, DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_LIGHT)


'''--------TransactionCode-/nMM43----------''''

Call SetTcode(DT_ZMDPC_UPLOAD_COND_0500_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Article","RMMW1-MATNR","",DT_ZMDPC_UPLOAD_COND_0100_ARTICLE_OCC1,false)
Call SetTextbox("Purchasing Org.","RMMW1-EKORG","",DT_ZMDPC_UPLOAD_COND_0100_PURCHASING_ORG_OCC1,false)
call SetTextboxNoLabel("RMMW1-LIFNR",2,DT_ZMDPC_UPLOAD_COND_0100_VENDOR,false)
'Call SetTextbox("Vendor","RMMW1-LIFNR","",DT_ZMDPC_UPLOAD_COND_0100_VENDOR_OCC1,false)  '' Sandbox changes from vendor to supplier
Call ClickButton("Deselect All   \(Shift\+F7\)",False)
Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100", "Screen description", "Purchasing", False)
Call PressEnter() 
Call ClickButton("For the Conditions",False)
Call TakeScreenShot
Call SelectCellGuiTable("SAPLV14ATCTRL_D0102", "Valid From", "Valid to", "31.03.2019", True)

Call ClickButton("btn\[0\]",True)
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "Supplier", 0, DT_ZMDPC_UPLOAD_COND_0010_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LIFNR)
'Call VerifyGridCellContent("", 1, "Vendor", 0, DT_ZMDPC_UPLOAD_COND_0010_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LIFNR_OCC1)  '' Sandbox changes from vendor to supplier
Call VerifyGridCellContent("", 1, "Article", 0, DT_ZMDPC_UPLOAD_COND_0010_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR_OCC1)
Call VerifyTableCellContent(2, "Amount", "SAPMV13ATCTRL_D0201", DT_ZMDPC_UPLOAD_COND_0201_CHECK_TEXT_OF_TABLECELL_AMOUNT_2_OCC1)
Call SelectMenuBar("Environment;Changes;Per Condition Record")
Call SetFocusGuiLabel(lcase(DT_ZMDPC_UPLOAD_COND_0120_CHECK_TEXT_OF_RECORD_DELETED_DUE_TO_OVERLAPPING_VALIDPERIOD), 67, 216, false)
Call VerifyifGuiLabelExists("Record deleted due to overlapping valid.period")

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


