
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AR0011 Manage AR Documents-Manage Manual Customer Invoicing_Cred 
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


gstrTestCaseName = "Test_AR0011 Manage AR Documents-Manage Manual Customer Invoicing_Cred"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\FICO\TASE_DT_AR0001 Manage AR Documents-Manage Manual Customer Invoicing_Cred.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


''''--------TransactionCode-FB70----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

''Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB70_1000_COMPANY_CODE,True)
Call SetTextboxPopupIfExist("BKPF-BUKRS", "Company Code", DT_FB70_1000_COMPANY_CODE)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call SetComboByKey("Transactn", DT_FB70_1200_TRANSACTN)
Call TakeScreenShot
Call SetTextbox("Customer","INVFO-ACCNT","",DT_FB70_0510_CUSTOMER,False)
Call PressEnter()
Call SetTextbox("Invoice date","INVFO-BLDAT","",ConvertDate(DT_FB70_0510_INVOICE_DATE),False)
Call SetTextbox("Posting Date","INVFO-BUDAT","",ConvertDate(DT_FB70_0510_POSTING_DATE),False)
Call PressEnter()
Call PressEnter()
Call SetTextbox("Reference","INVFO-XBLNR","",DT_FB70_0510_REFERENCE,False)
Call PressEnter()
Call PressEnter()
Call SetTextbox("HeaderText","INVFO-BKTXT","",DT_FB70_0550_HEADERTEXT,False)
Call SelectTab("TS", "Basic data", blnIsItPopup)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB70_0510_AMOUNT,False)
Call SetTextbox("Amount","INVFO-WAERS","",DT_FB70_0510_AMOUNT_OCC1,False)
Call SetTextbox("Tax Amount","INVFO-WMWST","",DT_FB70_0510_TAX_AMOUNT,False)
Call SetTextbox("Text","INVFO-SGTXT","",DT_FB70_0510_TEXT,False)
Call SetTableData("SAPLFSKBTABLE", "G/L acct", 1, "", "", DT_FB70_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", 1, "", "", DT_FB70_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "Tax code", 1, "", "", DT_FB70_0100_TABLECELL_TAX_CODE_0, False)
Call SetTableData("SAPLFSKBTABLE", "Business area", 1, "", "", DT_FB70_0100_TABLECELL_BUSINESS_AREA_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", 1, "", "", DT_FB70_0100_TABLECELL_COST_CENTER_0, False)
Call TakeScreenShot
Call PressEnter()
Call PressEnter()
Call TakeScreenShot
Call PressEnter()
Call SelectMenuBar("Document;Simulate")
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)", False)
Call GetStatusBar("item1", "DT_INVOICE_NO_OUTPUT")
Call VerifyStatusBar("Document "&DT_INVOICE_NO_OUTPUT&" was posted in company code RO02")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",4)
Call ClickButtonIfExist("Cancel   \(F12\)", True)

'''''''--------TransactionCode-/FB03----------''''

Call SetTcode(DT_FB70_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call SetTextbox("Document Number","RF05L-BELNR","",DT_INVOICE_NO_OUTPUT,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB70_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FB70_0100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()
'Call ClickButtonToolBar("&FIND", 0)
'Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB70_0841_SEARCH_TERM,True)
'Call TakeScreenShot
'Call ClickButtonIfExist("OK   \(Enter\)", True)
'Call ClickButtonIfExist("Cancel   \(F12\)", True)
Call GetTextboxValue("BKPF-XBLNR", "", "DT_FB70_0750_GET_TEXT_OF_REFERENC_OUTPUT", False)
Call VerifyTextBoxContent("Currency","BKPF-WAERS",0,DT_FB70_0750_CHECK_TEXT_OF_CURRENCY, False)
Call VerifyGridCellContent("", 1, "Assignment", 0, DT_FB70_0750_GET_TEXT_OF_REFERENC_OUTPUT)
Call VerifyGridCellContent("", 1, "Posting Key", 0, DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "Posting Key", 0, DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 3, "Posting Key", 0, DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call DoubleClickGuiGridCell("",0, 1, "Company code", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Acc","BSEG-HKONT",0,DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT, False)
Call VerifyTextBoxContent("Bline Date","BSEG-ZFBDT",0,DT_FB70_0301_CHECK_TEXT_OF_BLINE_DATE, False)
Call ClickButton("Back   \(F3\)", False)
Call DoubleClickGuiGridCell("",0, 2, "Posting Key", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Account","BSEG-HKONT",0,DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT, False)
Call ClickButton("Back   \(F3\)", False)
Call DoubleClickGuiGridCell("",0, 3, "Posting Key", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Account","BSEG-HKONT",0,DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_HKONT, False)
Call ClickButton("Back   \(F3\)", False)
Call VerifyGridCellContent("", 1, "Account", 0, DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "Account", 0, DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 3, "Account", 0, DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
Call VerifyGridCellContent("", 1, "AZBET", 0, DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("", 2, "AZBET", 0, DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("", 3, "AZBET", 0, DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AZBET)
'Call VerifyGridCellContent("", 3, "RF05A_UBAZW", 0, DT_CURR)
Call ClickButton("Back   \(F3\)", False)
Call ClickButton("Back   \(F3\)", False)


''''''--------TransactionCode-/nfbl5n----------''''

Call SetTcode(DT_FB70_0100_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)

Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_FB70_1000_CUSTOMER_ACCOUNT,False)
Call SelectRadioButton("X_AISEL", "All items", False)
Call ClickButton("Dynamic selections   \(Shift\+F4\)", False)
Call ClickButton("%_%%DYN011_%_APP_%-VALU_PUSH", False)
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "", "", DT_INVOICE_NO_OUTPUT, True)
Call ClickButton("Copy   \(F8\)", True)
Call ClickButton("Execute   \(F8\)", False)
Call TakeScreenShot
Call VerifyifGuiLabelExists_ByIndex(DT_FB70_0750_GET_TEXT_OF_REFERENC_OUTPUT,0)
Call VerifyifGuiLabelExists_ByIndex(DT_INVOICE_NO_OUTPUT,0)
Call ClickButton("Change layout...   \(Ctrl\+F8\)", False)
Call SelectRowGuiTableByRow("SAPLSKBHTC_FIELD_LIST",1, True)
Call ClickButton("Show Sel. Fields \(Ctrl\+F3\)", True)
Call SelectRowGuiTableByRow("SAPLSKBHTC_FIELD_LIST",72, True)
Call ClickButton("Show Sel. Fields \(Ctrl\+F3\)", True)
Call SelectRowGuiTableByRow("SAPLSKBHTC_FIELD_LIST",18, True)
Call ClickButton("Show Sel. Fields \(Ctrl\+F3\)", True)
Call TakeScreenShot
Call ClickButton("Copy   \(Enter\)", True)
Call VerifyifGuiLabelExists_ByIndex(DT_FB70_0120_CHECK_TEXT_OF_DR,0)
Call VerifyifGuiLabelExists_ByIndex(DT_FB70_0120_CHECK_TEXT_OF_RON,0)
Call VerifyifGuiLabelExists_ByIndex("S_LEDR",0)
Call VerifyifGuiLabelExists_ByIndex(DT_FB70_0120_CHECK_TEXT_OF_NO_NAME_OCC1,0) ' amount in local currency
Call SetHorizontalScrollBar(50,False)
Call VerifyifGuiLabelExists_ByIndex(ConvertDate(DT_NET_DUE_DATE),0)
Call SelectMenuBar("Environment;Correspondence;Request")
Call TakeScreenShot
Call ClickButton("Restrict Values   \(Shift\+F5\)", True)
Call SetTextbox("Correspondence","G_SELFLD_TAB-LOW","","ZRO19",True)
Call ClickButton("Continue   \(Enter\)", True)
Call ClickButton("Copy   \(Enter\)", True)
Call SetTextbox("Document Number","RF022-BELNR","",DT_FB70_1001_DOCUMENT_NUMBER,True)
Call SetTextbox("Company Code","RF022-BUKRS","",DT_FB70_1001_COMPANY_CODE,True)
Call SetTextbox("Fiscal Year","RF022-GJAHR","",DT_FB70_1001_FISCAL_YEAR,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)", True)
Call VerifyStatusBar(DT_FB70_0120_CHECK_TEXT_OF_STATUSBAR)
Call SelectMenuBar("Environment;Correspondence;Display")
Call TakeScreenShot
Call SetTextbox("Output Device","USR01-SPLD","",DT_FB70_1100_OUTPUT_DEVICE,True)
Call PressEnter()
Call ClickButtonIfExist("Yes",True)
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




