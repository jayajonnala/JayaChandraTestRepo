
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AR0002 Manage AR Documents-Manage Manual Customer Invoicing_Cred
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


gstrTestCaseName = "Test_AR0002 Manage AR Documents-Manage Manual Customer Invoicing_Cred"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\FICO\TASE_DT_AR0002 Manage AR Documents-Manage Manual Customer Invoicing_Cred.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
''''''--------------login----------------'''''
'
'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


''''''--------TransactionCode-FB70----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

''Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB70_1000_COMPANY_CODE,True)
Call SetTextboxPopupIfExist("BKPF-BUKRS", "Company Code", DT_FB70_1000_COMPANY_CODE)
Call ClickButtonIfExist("Continue   \(Enter\)", True)

Call TakeScreenShot

Call SetComboByKey("Transactn", DT_FB70_1200_TRANSACTN)
Call TakeScreenShot
Call SetTextbox("Customer","INVFO-ACCNT","",DT_FB70_0510_CUSTOMER,False)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("Invoice date","INVFO-BLDAT","",ConvertDate(DT_FB70_0510_INVOICE_DATE),False)
Call SetTextbox("Posting Date","INVFO-BUDAT","",ConvertDate(DT_FB70_0510_POSTING_DATE),False)
Call TakeScreenShot
Call PressEnter()

Call PressEnter()

Call SetTextbox("Reference","INVFO-XBLNR","",DT_FB70_0510_REFERENCE,False)
Call TakeScreenShot
Call PressEnter()
Call PressEnter()
Call SetTextbox("HeaderText","INVFO-BKTXT","",DT_FB70_0550_HEADERTEXT,False)
Call TakeScreenShot
Call SelectTab("TS", "Basic data", False)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB70_0510_AMOUNT,False)
Call PressEnter()
Call PressEnter()
Call SelectCheckbox("INVFO-XMWST", 0, DT_FB70_0510_CALCULATE_TAX, False)
Call TakeScreenShot
Call SetComboByKey("INVFO-MWSKZ", DT_FB70_0510_INVFOMWSKZ)
Call TakeScreenShot
Call SetTableData("SAPLFSKBTABLE", "G/L acct", 1, "", "", DT_FB70_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", 1, "", "", DT_FB70_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "Business area", 1, "", "", DT_FB70_0100_TABLECELL_BUSINESS_AREA_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", 1, "", "", DT_FB70_0100_TABLECELL_COST_CENTER_0, False)
Call TakeScreenShot
Call SelectTab("TS", "Payment", False)
Call TakeScreenShot
Call PressEnter()
Call SelectMenuBar("Document;Simulate")
Call TakeScreenShot
Call PressEnter()
Call VerifyifGuiLabelExists(DT_FB70_0120_CHECK_TEXT_OF_NO_NAME)
Call VerifyifGuiLabelExists(DT_FB70_0120_CHECK_TEXT_OF_NO_NAME_OCC1)
Call VerifyifGuiLabelExists(DT_FB70_0120_CHECK_TEXT_OF_NO_NAME_OCC2)
Call TakeScreenShot
Call PressEnter()
Call ClickButton("Post   \(Ctrl\+S\)", False)
Call GetStatusBar("item1", "DT_DOC_NO_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_NO_OUTPUT&" was posted in company code RO02")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",4)
Call ClickButtonIfExist("Cancel   \(F12\)", True)

'''''''--------TransactionCode-/FB03----------''''

Call SetTcode(DT_FB70_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

'Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB70_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB70_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FB70_0100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()
Call GetGridContent("", 0, "ZUONR", 1, "<NA>", "<NA>", "DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_OUTPUT")
Call VerifyTextBoxContent("Reference","BKPF-XBLNR","",DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_OUTPUT, False)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 3, "KTONR", 0, DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_PK_1)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_PK_2)
Call VerifyGridCellContent("", 3, "BSCHL", 0, DT_PK_3)
Call TakeScreenShot
''Call ClickButtonToolBar("&FIND", 0)
''Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB70_0841_SEARCH_TERM,True)
'Call TakeScreenShot
'Call ClickButtonIfExist("OK   \(Enter\)", True)
'Call ClickButtonIfExist("Cancel   \(F12\)", True)
Call DoubleClickGuiGridCell("",0, 1, "Company code", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Acc","BSEG-HKONT",0,DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT, False)
Call ClickButton("Back   \(F3\)", False)
'Call VerifyGridCellContent("", 1, "G/L Account", 0, DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
Call VerifyGridCellContent("", 2, "Account", 0, DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)
Call VerifyGridCellContent("", 3, "Account", 0, DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_HKONT)
Call VerifyGridCellContent("", 1, "Amount", 0, DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("", 2, "Amount", 0, DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("", 3, "Amount", 0, DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AZBET)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)", False)
Call ClickButton("Back   \(F3\)", False)

''''''''--------TransactionCode-/nfbl5n----------''''

Call SetTcode(DT_FB70_0100_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)

Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_FB70_1000_CUSTOMER_ACCOUNT,False)
Call SelectRadioButton("X_AISEL", "All items", False)
Call ClickButton("Dynamic selections   \(Shift\+F4\)", False)
Call ClickButton("%_%%DYN011_%_APP_%-VALU_PUSH", False)
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "", "", DT_DOC_NO_OUTPUT, True)
Call ClickButton("Copy   \(F8\)", True)
Call ClickButton("Execute   \(F8\)", False)
Call VerifyifGuiLabelExists_ByIndex(DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_OUTPUT,0)
'Call VerifyifGuiLabelExists(DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_OUTPUT)
Call ClickButton("Back   \(F3\)", False)
Call ClickButton("Back   \(F3\)", False)

'''''''''--------TransactionCode-/F.62----------''''

Call SetTcode(DT_FB70_0100_OKCD_OCC2)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC6)

Call SelectCheckbox("NORMBL", 0, DT_FB70_1000_STANDARD_DOCUMENTS, False)
Call SetTextbox("Fiscal year","RGJAHR-LOW","",DT_FB70_1000_FISCAL_YEAR,False)
Call SetTextbox("Document number","RBELNR-LOW","",DT_DOC_NO_OUTPUT,False)
Call SetTextbox("Company code","RBUKRS-LOW","",DT_FB70_1000_COMPANY_CODE_OCC1,False)
Call SetTextbox("Document type","RBLART-LOW","",DT_FB70_1000_DOCUMENT_TYPE,False)
Call SetTextbox("Correspondence","REVENT","",DT_FB70_1000_CORRESPONDENCE,False)
Call SetTextbox("Posting date","RBUDAT-LOW","",ConvertDate(DT_FB70_1000_POSTING_DATE),False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)", False)
Call ClickButtonIfExist("Continue   \(Enter\)", True)
Call ClickButtonIfExist("Yes", True)
Call ClickButtonIfExist("Yes", True)
Call ClickButtonIfExist("Yes", True)
Call SetTextbox("Output Device","USR01-SPLD","",DT_FB70_1100_OUTPUT_DEVICE,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)", True)
Call ClickButtonIfExist("Yes", True)
Call ClickButton("Back   \(F3\)", False)
Call TakeScreenShot
Call VerifyifGuiLabelExists_ByIndex(DT_FB70_0120_CHECK_TEXT_OF_RO02,0)
Call VerifyifGuiLabelExists_ByIndex(lcase(DT_FB70_0120_CHECK_TEXT_OF_CUSTOMER_INVOICE__MI),0)
Call ClickLabel("RO02", "", False)
Call ClickButton("Details   \(Ctrl\+Shift\+F3\)", False)
Call VerifyifGuiLabelExists(lcase(DT_FB70_0120_CHECK_TEXT_OF_ZRO19))
Call ClickButton("Back   \(F3\)", False)
Call ClickButton("Back   \(F3\)", False)

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




