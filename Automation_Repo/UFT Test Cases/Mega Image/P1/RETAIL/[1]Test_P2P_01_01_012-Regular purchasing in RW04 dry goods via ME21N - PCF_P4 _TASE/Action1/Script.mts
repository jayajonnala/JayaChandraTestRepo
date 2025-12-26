
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_012-Regular purchasing in RW04 dry goods via ME21N - PCF_P4
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

gstrTestCaseName = "Test_P2P_01_01_012-RW04 dry goods via ME21N - PCF_P4"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_P2P_01_01_012-Regular purchasing in RW04  dry goods  via ME21N - PCF_P4.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
'
'
''''--------TransactionCode-ME23N----------''''
'
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot

Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call SetTextbox("Pur. Order","MEPO_SELECT-EBELN","",DT_ME23N_0003_PUR_ORDER,True) 
Call SelectRadioButton("MEPO_SELECT-BSTYP_F","Pur. Order",True) 
Call ClickButton("Other Document   \(Enter\)",True)
Call ClickButtonIfExist("Expand Item Details Ctrl\+F4",False)
Call SelectTab("ITEM_DETAIL", "Purchase Order History", False)
Call TakeScreenShot
Call GetGridContent("", 0, "BELNR", 1, "<NA>", "<NA>", "DT_GET_ARTICLE_DOCUMENT_NUMBER_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_GET_ARTICLE_DOCUMENT_NUMBER_OUTPUT",DT_GET_ARTICLE_DOCUMENT_NUMBER)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call TakeScreenShot


'''--------TransactionCode-MIGO----------''''

Call SetTcode(DT_ME23N_0014_OKCD)     
Call PressEnter()     
Call TakeScreenShot

''Call SendKey("^{F3}")
''Call TakeScreenShot
''Call SendKey("^{F4}")
''Call TakeScreenShot
Call SetCombo("GODYNPRO-ACTION", "Display")

Call SetTextboxNoLabel("GODYNPRO-MAT_DOC", 0, DT_ME23N_2010_GODYNPROMAT_DOC, false)
Call SetTextboxNoLabel("GODYNPRO-DOC_YEAR", 0, DT_ARTICLEDOC_YEAR, false)
Call TakeScreenShot
Call PressEnter()

''Call GetTableCellData("SAPLMIGOTV_GOITEM",  "Art. Short Text" , 1, "<NA>", "<NA>", "DT_ARTICLE_NAME_1_OUTPUT", False)
Call GetTableCellData("SAPLMIGOTV_GOITEM",  "Mat. Short Text" , 1, "<NA>", "<NA>", "DT_ARTICLE_NAME_1_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_ARTICLE_NAME_1_OUTPUT",DT_ARTICLE_NAME_1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

''Call GetTableCellData("SAPLMIGOTV_GOITEM",  "Art. Short Text" , 2, "<NA>", "<NA>", "DT_ARTICLE_NAME_2_OUTPUT", False)
Call GetTableCellData("SAPLMIGOTV_GOITEM",  "Mat. Short Text" , 2, "<NA>", "<NA>", "DT_ARTICLE_NAME_2_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_ARTICLE_NAME_2_OUTPUT",DT_ARTICLE_NAME_2)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'Call VerifyTableCellContent(1, "Quantity", "SAPLMIGOTV_GOITEM", DT_ARTICLE_1_QUANTITY)
Call VerifyTableCellContent(2, "Quantity", "SAPLMIGOTV_GOITEM", DT_ARTICLE_2_QUANTITY)
Call VerifyTableCellContent(1, "Movement type", "SAPLMIGOTV_GOITEM", DT_ARTICLE_1_MOVEMENT_TYPE)
Call VerifyTableCellContent(2, "Movement type", "SAPLMIGOTV_GOITEM", DT_ARTICLE_2_MOVEMENT_TYPE)

Call SelectTab("TS_GOITEM","Output",False)
Call TakeScreenShot
Call ClickButton("Display outputs",false)
Call TakeScreenShot

Call VerifyTableCellContent(2, "Output Type", "SAPDV70ATC_NAST3", DT_ME23N_0100_CHECK_ACCTEXT_OF_TABLECELL_OUTPUT_TYPE_0)
Call VerifyTableCellContent(2, "Status", "SAPDV70ATC_NAST3", DT_ME23N_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_0)
Call VerifyTableCellContent(3, "Output Type", "SAPDV70ATC_NAST3", DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_1)
Call VerifyTableCellContent(3, "Status", "SAPDV70ATC_NAST3", DT_ME23N_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_1)
Call TakeScreenShot

Call ClickButton("Back   \(F3\)",false)
Call SelectTab("TS_GOHEAD","Doc. info",False)
Call ClickButton("FI Documents",false)
'Call SelectRowGuiGrid("Documents in Accounting", 0, "Object type text", "Accounting document", True)
'Call ClickButtonIfExist("Display Document   \(F2\)",True)
Call GetTextboxValue("BKPF-BELNR", 0, "DT_GET_ACCOUNTING_DOCUMENT_NUMBER_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_GET_ACCOUNTING_DOCUMENT_NUMBER_OUTPUT",DT_GET_ACCOUNTING_DOCUMENT_NUMBER)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_ME23N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_ME23N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_ME23N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_ME23N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call TakeScreenShot

''''''--------TransactionCode-MB90----------''''
'
Call SetTcode(DT_ME23N_0100_OKCD)     
Call PressEnter()     

Call SetTextbox("Article Document","RG_MBLNR-LOW","",DT_ME23N_1000_ARTICLE_DOCUMENT,False)
Call SetTextbox("Article Doc. Year","PM_MJAHR","",DT_ARTICLEDOC_YEAR,False)
Call SetTextbox("Processing Mode","PM_VERMO","",DT_ME23N_1000_PROCESSING_MODE,False)
Call SetTextbox("Sort order","PM_NSORT","",DT_ME23N_1000_SORT_ORDER,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call SelectCheckboxNoLabel(0, "ON", False)
Call ClickButton("Print preview   \(Shift\+F4\)",False)
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




