
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_01-Regular purchasing in RW04 dry goods via ME21N - P&Z_P4
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



gstrTestCaseName = "Test_P2P_01_01_01  RW04 dry goods via ME21N - P&Z_P4"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_P2P_01_01_01-Regular purchasing in RW04  dry goods  via ME21N - P&Z_P4.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''---------Login-------------''''''
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'''''''-------Transaction Code ME21N-------'''''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter() 
Call TakeScreenShot

Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call SetTextbox("Pur. Order","MEPO_SELECT-EBELN","",DT_ME23N_0003_PUR_ORDER,True) 
Call SelectRadioButton("MEPO_SELECT-BSTYP_F","Pur. Order",True) 
Call ClickButton("Other Document   \(Enter\)",True)
Call ClickButtonIfExist("Expand Header Ctrl\+F2",false)
Call SelectTab("HEADER_DETAIL","Texts",False)
Call TakeScreenShot
Call SelectTab("HEADER_DETAIL","Org. Data",False)
Call TakeScreenShot
Call ClickButtonIfExist("Expand Item Details Ctrl\+F4",False)
Call SelectTab("ITEM_DETAIL","Purchase Order History",False)
Call TakeScreenShot
Call GetGridContent("", "", "Article Document", 1, "<NA>", "<NA>", "DT_GET_ARTICLE_DOCUMENT_NUMBER_OUTPUT")
Call VerifyGridCellContentbyName("shell", 1, "Article Document", "", DT_GET_ARTICLE_DOCUMENT_NUMBER_OUTPUT)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_GET_ARTICLE_DOCUMENT_NUMBER_OUTPUT",DT_GET_ARTICLE_DOCUMENT_NUMBER)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call TakeScreenShot

'''-------TransactionCode---MIGO--------'''

Call SetTcode(DT_ME23N_0014_OKCD)     
Call PressEnter()     

Call SetComboByKey("GODYNPRO-ACTION", DT_ME23N_0010_GODYNPROACTION)
Call SetTextboxNoLabel("GODYNPRO-MAT_DOC", "", DT_ME23N_2010_GODYNPROMAT_DOC, False)
Call PressEnter()
''Call GetTableCellData("SAPLMIGOTV_GOITEM", "Art. Short Text", 1, "<NA>", "<NA>", "DT_ARTICLE_NAME_1_OUTPUT", False)
Call GetTableCellData("SAPLMIGOTV_GOITEM", "Mat. Short Text", 1, "<NA>", "<NA>", "DT_ARTICLE_NAME_1_OUTPUT", False)
Call VerifyTableCellContent(1, "Quantity", "SAPLMIGOTV_GOITEM", DT_ARTICLE_1_QUANTITY)
Call VerifyTableCellContent(1, "Movement type", "SAPLMIGOTV_GOITEM", DT_ARTICLE_1_MOVEMENT_TYPE)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_ARTICLE_NAME_1_OUTPUT",DT_ARTICLE_NAME_1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

''Call GetTableCellData("SAPLMIGOTV_GOITEM", "Art. Short Text", 1, "<NA>", "<NA>", "DT_ARTICLE_NAME_2_OUTPUT", False)
Call GetTableCellData("SAPLMIGOTV_GOITEM", "Mat. Short Text", 1, "<NA>", "<NA>", "DT_ARTICLE_NAME_2_OUTPUT", False)
Call VerifyTableCellContent(2, "Quantity", "SAPLMIGOTV_GOITEM", DT_ARTICLE_2_QUANTITY)
Call VerifyTableCellContent(2, "Movement type", "SAPLMIGOTV_GOITEM", DT_ARTICLE_2_MOVEMENT_TYPE)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_ARTICLE_NAME_2_OUTPUT",DT_ARTICLE_NAME_2)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'Call GetTableCellData("SAPLMIGOTV_GOITEM", "Art. Short Text", 1, "<NA>", "<NA>", "DT_ARTICLE_NAME_3_OUTPUT", False)
'Call VerifyTableCellContent(3, "Quantity", "SAPLMIGOTV_GOITEM", DT_ARTICLE_3_QUANTITY)
'Call VerifyTableCellContent(3, "Movement type", "SAPLMIGOTV_GOITEM", DT_ARTICLE_3_MOVEMENT_TYPE)
'Call WriteRunTimeDataToExcelGlobalSheet ("DT_ARTICLE_NAME_3_OUTPUT",DT_ARTICLE_NAME_3)
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
Call SelectTab("TS_GOITEM","Output",False)
Call TakeScreenShot
Call ClickButton("Display outputs",False)
Call TakeScreenShot
Call VerifyTableCellContent(2, "Status", "SAPDV70ATC_NAST3", DT_ME23N_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_0)
Call VerifyTableCellContent(2, "Output Type", "SAPDV70ATC_NAST3", DT_ME23N_0100_CHECK_ACCTEXT_OF_TABLECELL_OUTPUT_TYPE_0)
Call VerifyTableCellContent(3, "Status", "SAPDV70ATC_NAST3", DT_ME23N_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_1)
Call VerifyTableCellContent(3, "Output Type", "SAPDV70ATC_NAST3", DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_1)
Call ClickButton("Back   \(F3\)",False)
Call SelectTab("TS_GOHEAD","Doc. info",False)
Call ClickButton("FI Documents",false)
Call TakeScreenShot
'Call SelectRowGuiGrid("Documents in Accounting", 0, "Object type text", "Accounting document", True)
'Call ClickButtonIfExist("Display Document   \(F2\)",True)
Call GetTextboxValue("BKPF-BELNR", "", "DT_GET_ACCOUNTING_DOCUMENT_NUMBER_OUTPUT", false)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_GET_ACCOUNTING_DOCUMENT_NUMBER_OUTPUT",DT_GET_ACCOUNTING_DOCUMENT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call VerifyGridCellContentbyName("shell", 1, "Posting Key", "", DT_ME23N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContentbyName("shell", 2, "Posting Key", "", DT_ME23N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
'Call VerifyGridCellContentbyName("shell", 5, "Posting Key", "", DT_ME23N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
'Call VerifyGridCellContentbyName("shell", 6, "Posting Key", "", DT_ME23N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BSCHL)
Call VerifyGridCellContentbyName("shell", 1, "Account", "", DT_ME23N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContentbyName("shell", 2, "Account", "", DT_ME23N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
'Call VerifyGridCellContentbyName("shell", 5, "Account", "", DT_ME23N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
'Call VerifyGridCellContentbyName("shell", 6, "Account", "", DT_ME23N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KTONR)
Call TakeScreenShot

''''-------TransactionCode---/nMB90--------'''

Call SetTcode(DT_ME23N_0100_OKCD)     
Call PressEnter()

Call SetTextbox("Article Document","RG_MBLNR-LOW","",DT_ME23N_1000_ARTICLE_DOCUMENT,False) 
Call SetTextbox("Article Doc. Year","PM_MJAHR","",DT_ME23N_1000_ARTICLE_DOC_YEAR,False)
Call SetTextbox("Processing Mode","PM_VERMO","",DT_ME23N_1000_PROCESSING_MODE,False)
Call SetTextbox("Sort order","PM_NSORT","",DT_ME23N_1000_SORT_ORDER,False)
Call SetTextbox("Transmission Medium","RG_NACHA-LOW","","",False)
Call SetTextbox("Output Type","RG_KSCHL-LOW ","","",False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call SelectCheckboxNoLabel(0,"ON",False)
Call ClickButton("Print preview   \(Shift\+F4\)",False)
Call TakeScreenShot
Call SetTcode(DT_ME23N_0100_OKCD_OCC1)     
Call PressEnter()


Call LogOff()

Call FinalStatus ()



'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [13,4062483]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//








