'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_ACACTREE01-Create Accrual Engine_TASE
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_ACACTREE01-Create Accrual Engine"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\DLL_P3\Data\TASE_DT_02-04-01-05-03-Create new assortment-BASA.xls"

If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	datatable_row= Parameter("datatable_row")	
End If
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''''--------TransactionCode-/faglb03----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)


Call SetTextbox("Account Number","RACCT-LOW","",DT_FAGLB03_1000_ACCOUNT_NUMBER,False)
Call SetTextbox("Company Code","RBUKRS-LOW","",DT_FAGLB03_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RYEAR","",DT_FAGLB03_1000_FISCAL_YEAR,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
'Need to write generic way 

Call DoubleClickGuiGridCell("",1,DT_CURR_CELL,"Period",False)
Call TakeScreenShot
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
''''Call ClickButton("&FIND",True)' ClickButtonToolBar(buttonName, toolbarIndex)
''Call ClickButtonToolBar("&Find",0)
Call ClickButton("Find",True)
Call SetTextbox("Find","GD_SEARCHSTR","","Document Number",True)
Call ClickButton("Continue   \(Enter\)",True)
Call ClickButton("Show sel. fields \(CTRL\+F3\)",True)
Call ClickButton("Copy   \(Enter\)",True)
Call ClickButton("%_%%DYN001_%_APP_%-VALU_PUSH",True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "", "",DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR, True)
Call TakeScreenShot
Call ClickButton("Copy   \(F8\)",True)
Call ClickButton("Execute   \(Enter\)",True)
Call ClickButton("Change Layout\.\.\.   \(Ctrl\+F8\)",False)
Call ClickButton("Find",True)
''Call ClickButtonToolBar("&Find",0)
Call SetTextbox("Find","GD_SEARCHSTR","","Document Number",True)
Call ClickButton("Continue   \(Enter\)",True)
Call ClickButton("Show Sel\. Fields \(Ctrl\+F3\)",True)
Call ClickButton("Copy   \(Enter\)",True)
Call VerifyifGuiLabelExists(DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
'Call VerifyifGuiLabelExists(ConvertDate(DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))
Call VerifyifGuiLabelExists(DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)

''--------TransactionCode--ACACTREE01---------''''

Call SetTcode(DT_FAGLB03_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call SetTextbox("Company Code","ACEDSOH-BUKRS","",DT_FAGLB03_1000_COMPANY_CODE_OCC1,False)
Call SetTextbox("Accrual Object Cat.","ACAC_OBJECTS-ACAC_OBJTYPE","",DT_FAGLB03_1000_ACCRUAL_OBJECT_CAT,False)
Call SetTextbox("End of Life","ACEDSOH-VALITY_TO","",ConvertDate(DT_FAGLB03_1000_END_OF_LIFE),False)
Call SetTextbox("Start of Life","ACEDSOH-VALITY_FROM","",ConvertDate(DT_FAGLB03_1000_START_OF_LIFE),False)
Call SetTextbox("Text","ACAC_OBJECTS-TEXT","",DT_FAGLB03_1000_TEXT,False)
Call TakeScreenShot
SAPGuiSession("Session").SAPGuiWindow("Create Accrual Object").InsightObject("InsightObject").Click
Call SetGridData("",1,"Accrual Type", DT_FAGLB03_3000_GRIDCELL_0_ACCRL_TYPE_OCC1,False)
Call SetGridData("",1,"Accounting Principle", DT_FAGLB03_3000_GRIDCELL_0_ACCP,False)
Call SetGridData("",1,"Total Amount To Be Accrued", DT_FAGLB03_3000_GRIDCELL_0_TOTAL_AMT_TO_BE_ACCD,False)
Call SetGridData("",1,"Accr. item currency",DT_FAGLB03_3000_GRIDCELL_0_ACRCY,False)
Call SetGridData("",1,"Total Qty to Accrue",DT_FAGLB03_3000_GRIDCELL_0_TOTAL_QTY_TO_ACCRUE ,False)
Call SetGridData("",1,"Acc./Defer. Method",DT_FAGLB03_3000_GRIDCELL_0_ACCRUAL_METH,False)
Call TakeScreenShot

Call SelectTab("TABSTRIP_POSITIONS", "Acct Assgts", False)
Call SetGridData("",1,"Business Area",DT_FAGLB03_0200_GRIDCELL_0_BUSA ,False)
Call SetGridData("",1,"Profit Center",DT_FAGLB03_0200_GRIDCELL_0_PROFIT_CENTER ,False)
Call SetGridData("",1,"Cost Center",DT_FAGLB03_0200_GRIDCELL_0_COST_CENTER,False)
Call TakeScreenShot

Call SelectTab("TABSTRIP_DETAILS", "Accts", False)
Call SetTextbox("Start Acct","ACEPS_ACCOUNTS_WITH_F4_R3-INC_START_ACCNT","",DT_FAGLB03_0100_START_ACCT,False)
Call SetTextbox("Targ. Acct","ACEPS_ACCOUNTS_WITH_F4_R3-INC_TARGET_ACCNT","",DT_FAGLB03_0100_TARG_ACCT,False)
Call SetTextbox("Document Type","ACEPS_ACCOUNTS_WITH_F4_R3-INC_DOC_TYPE","",DT_FAGLB03_0100_DOCUMENT_TYPE,False)
Call SetTextbox("Start Acct","ACEPS_ACCOUNTS_WITH_F4_R3-PER_START_ACCNT","",DT_FAGLB03_0100_START_ACCT_OCC1,False)
Call SetTextbox("Targ. Acct","ACEPS_ACCOUNTS_WITH_F4_R3-PER_TARGET_ACCNT","",DT_FAGLB03_0100_TARG_ACCT_OCC1,False)
Call SetTextbox("Document Type","ACEPS_ACCOUNTS_WITH_F4_R3-PER_DOC_TYPE","",DT_FAGLB03_0100_DOCUMENT_TYPE_OCC1,False)  
Call SetTextbox("Start Acct","ACEPS_ACCOUNTS_WITH_F4_R3-FIN_START_ACCNT","",DT_FAGLB03_0100_START_ACCT_OCC2,False)
Call SetTextbox("Targ. Acct","ACEPS_ACCOUNTS_WITH_F4_R3-FIN_TARGET_ACCNT","",DT_FAGLB03_0100_TARG_ACCT_OCC2,False)
Call SetTextbox("Document Type","ACEPS_ACCOUNTS_WITH_F4_R3-FIN_DOC_TYPE","",DT_FAGLB03_0100_DOCUMENT_TYPE_OCC2,False) 
Call TakeScreenShot

Call ClickButton("Check and Simulate Changes   \(Ctrl\+Shift\+F2\)",False)
Call SelectRadioButton("ACE_REP_PARAMETER-FLAG_MAN_PERIOD_POST", "Period Entered Manually", True)
Call SetTextbox("Fiscal Year","ACEPS_MANUAL_POSTING_PARAMS-FISC_YEAR","",DT_FAGLB03_0100_FISCAL_YEAR,True)
Call SetTextbox("Posting Period","ACEPS_MANUAL_POSTING_PARAMS-FIS_PERIOD","",DT_FAGLB03_0100_POSTING_PERIOD,True)
Call TakeScreenShot
Call ClickButtonIfExist("Copy the Parameters Entered   \(Enter\)",True)
Call TakeScreenShot

Call VerifyifGuiLabelExists(DT_FAGLB03_0120_CHECK_TEXT_OF_TEST_RUN_NO_DATA_UPDATED)
'Call VerifyifGuiLabelExists(DT_FAGLB03_0120_CHECK_OCC1)
Call VerifyifGuiLabelExists(DT_FAGLB03_0120_CHECK_OCC2)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)

Call SelectRadioButton("ACE_REP_PARAMETER-FLAG_MAN_PERIOD_POST", "Period Entered Manually", True)
Call SetTextbox("Fiscal Year","ACEPS_MANUAL_POSTING_PARAMS-FISC_YEAR","",DT_FAGLB03_0100_FISCAL_YEAR,True)
Call SetTextbox("Posting Period","ACEPS_MANUAL_POSTING_PARAMS-FIS_PERIOD","",DT_FAGLB03_0100_POSTING_PERIOD,True)
Call TakeScreenShot
Call ClickButtonIfExist("Copy the Parameters Entered   \(Enter\)",True)
Call ClickButtonIfExist("Current display variant   \(Ctrl\+F8\)",True)
Call SetTableData("SAPLSKBHTC_WRITE_LIST","Length",2, "", "", 72, True)
Call ClickButtonIfExist("Copy   \(Enter\)",True)

Call GetLabelContentByRefLabel("Message text",0,-32,"DT_OP_FAGLB03_0120_CHECK_TEXT_OF_DOCUMENT_POSTED_SUCCESSFULLY_ACE_2000001900BE102015_R1ECLNT200",True)
Call GetLabelContentByRefLabel("Message text",0,-48,"DT_OP_FAGLB03_0120_CHECK_TEXT_OF_ACCRUAL_OBJECT_WITH_EXTERNAL_REFERENCE_D08_SERVI_00000000000000000101",True)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",4)

Call VerifyifGuiLabelExists(DT_FAGLB03_0120_CHECK_TEXT_OCC1)
Call VerifyifGuiLabelExists(DT_FAGLB03_0120_CHECK_TEXT_OCC2)
Call ClickButtonIfExist("Continue   \(Enter\)",True)

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

