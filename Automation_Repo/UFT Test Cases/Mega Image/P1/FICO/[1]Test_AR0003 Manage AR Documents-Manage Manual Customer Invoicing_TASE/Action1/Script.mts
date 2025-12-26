			
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AR0003 Manage AR Documents-Manage Manual Customer Invoicing
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


gstrTestCaseName = "Test_AR0003 Manage AR Documents-Manage Manual Customer Invoicing"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\FICO\TASE_DT_AR0003 Manage AR Documents-Manage Manual Customer Invoicing.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
''''''--------------login----------------'''''
'''
''''''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
'''''
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'
'''''--------TransactionCode--FD10N----------''''
'
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",DT_INCREMENT+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextbox("Customer","SO_KUNNR-LOW","",DT_FD10N_1000_CUSTOMER,False)
Call SetTextbox("Company code","SO_BUKRS-LOW","",DT_FD10N_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal year","GP_GJAHR","",DT_FD10N_1000_FISCAL_YEAR,False)
Call PressEnter()
Call ClickButton("Execute   \(F8\)", False)

Call GetGridContent("", 0, "Balance", 18, "<NA>", "<NA>", "DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_OUTPUT",DT_FD10N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_17_BALANCE)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyGridCellContent("", 18, "Cumulative balance", 0, "")

Call TakeScreenShot
Call ClickButton("Back   \(F3\)", False)
Call ClickButton("Back   \(F3\)", False)

'''''''--------TransactionCode-/FB70----------''''

Call SetTcode(DT_FD10N_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

''Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FD10N_1000_COMPANY_CODE_OCC1,True)
''Call TakeScreenShot
''Call PressEnter()
Call SetTextboxPopupIfExist("BKPF-BUKRS", "Company Code", DT_FD10N_1000_COMPANY_CODE_OCC1)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)", True)
Call SetComboByKey("Transactn", DT_FD10N_1200_TRANSACTN)
Call TakeScreenShot
Call SetTextbox("Customer","INVFO-ACCNT","",DT_FD10N_0510_CUSTOMER,False)
Call SetTextbox("Invoice date","INVFO-BLDAT","",ConvertDate(DT_FD10N_0510_INVOICE_DATE),False)
Call SetTextbox("Reference","INVFO-XBLNR","",DT_FD10N_0510_REFERENCE,False)
Call SetTextbox("Posting Date","INVFO-BUDAT","",ConvertDate(DT_FD10N_0510_POSTING_DATE),False)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FD10N_0510_AMOUNT,False)
Call SetTableData("SAPLFSKBTABLE", "G/L acct", 1, "", "", DT_FD10N_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", 1, "", "", DT_FD10N_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "Tax code", 1, "", "", DT_FD10N_0100_TABLECELL_TAX_CODE_0, False)
Call SetTableData("SAPLFSKBTABLE", "Business area", 1, "", "", DT_FD10N_0100_TABLECELL_BUSINESS_AREA_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", 1, "", "", DT_FD10N_0100_TABLECELL_COST_CENTER_0, False)
Call SelectTab("TS", "Details", False)
Call PressEnter()
Call PressEnter()

Call PressEnter()

Call SetTextbox("HeaderText","INVFO-BKTXT","",DT_FD10N_0550_HEADERTEXT,False)
Call TakeScreenShot
Call SelectTab("TS", "Basic data", False)
Call PressEnter()
Call SetTextbox("Text","INVFO-SGTXT","",DT_FD10N_0510_TEXT,False)
Call ClickButton("Simulate Document Posting   \(F9\)", False)
Call PressEnter()
Call SelectCheckbox("INVFO-XMWST", 0, DT_FD10N_0510_CALCULATE_TAX, False)
Call ClickButton("Simulate Document Posting   \(F9\)", False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)", False)

Call GetStatusBar("item1", "DT_DOC_NO_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_DOC_NO_OUTPUT",DT_DOC_NO)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'Call VerifyStatusBar("Document "&DT_DOC_NO_OUTPUT&" was posted in company code RO02")
Call VerifyStatusBar(DT_FD10N_1000_CHECK_TEXT_OF_STATUS_BAR)
Call ClickButtonIfExist("Cancel   \(F12\)", True)

''''''''--------TransactionCode-/FB03----------''''

Call SetTcode(DT_FD10N_0100_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Document Number","RF05L-BELNR","",DT_DOC_NO,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FD10N_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FD10N_0100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()

Call GetTextboxValue("BKPF-XBLNR", "", "DT_FD10N_0750_CHECK_TEXT_OF_REFERENCE_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_FD10N_0750_CHECK_TEXT_OF_REFERENCE_OUTPUT",DT_FD10N_0750_CHECK_TEXT_OF_REFERENCE)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call VerifyGridCellContent("", 1, "ZUONR", 0, DT_FD10N_0750_CHECK_TEXT_OF_REFERENCE)
Call VerifyGridCellContent("", 1, "Posting Key", 0, DT_FD10N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "Posting Key", 0, DT_FD10N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 3, "Posting Key", 0, DT_FD10N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("", 1, "Account", 0, DT_FD10N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "Account", 0, DT_FD10N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 3, "Account", 0, DT_FD10N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
Call VerifyGridCellContent("", 1, "AZBET", 0, DT_FD10N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("", 2, "AZBET", 0, DT_FD10N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("", 3, "AZBET", 0, DT_FD10N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AZBET)
Call TakeScreenShot
Call SelectMenuBar("Environment;Correspondence")
Call ClickButton("Find   \(Ctrl\+F\)", True)
Call SetTextbox("Find","RSYSF-STRING","",DT_FD10N_0800_FIND,True)
Call TakeScreenShot
Call ClickButtonIfExist("Find   \(Enter\)", True)
Call SetFocusGuiLabel("Customer Invoice - MI", 53, 40, True)
Call SendKey("{F2}")
Call ClickButtonIfExist("Copy   \(Enter\)", True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)", True)
Call VerifyStatusBar(DT_FD10N_0750_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)", False)
Call ClickButton("Back   \(F3\)", False)

'''''''--------TransactionCode-/nfbl5n----------''''

Call SetTcode(DT_FD10N_0100_OKCD_OCC2)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC6)

Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_FD10N_1000_CUSTOMER_ACCOUNT,False)
Call SelectRadioButton("X_AISEL", "All items", False)
Call ClickButton("Dynamic selections   \(Shift\+F4\)", False)
Call ClickButton("%_%%DYN011_%_APP_%-VALU_PUSH", False)
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "", "", DT_DOC_NO, True)
Call ClickButton("Copy   \(F8\)", True)
Call ClickButton("Execute   \(F8\)", False)
Call VerifyifGuiLabelExists_ByIndex(DT_FD10N_0750_CHECK_TEXT_OF_REFERENCE_OCC1,0)
Call VerifyifGuiLabelExists_ByIndex(DT_DOC_NO,0)
Call VerifyifGuiLabelExists_ByIndex(DT_FD10N_0120_CHECK_TEXT_OF_DR,0)
'Call VerifyifGuiLabelExists(DT_FD10N_0120_CHECK_TEXT_OF_NO_NAME_OCC1)
'Call VerifyifGuiLabelExists(DT_FD10N_0120_CHECK_TOOLTIP_OF_NO_NAME)
Call VerifyifGuiLabelExists_ByIndex(DT_FBL5N_0120_CHECK_TEXT_OF_NO_NAME,0)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)", False)
Call ClickButton("Back   \(F3\)", False)

'''''''''--------TransactionCode-/F.62----------''''

Call SetTcode(DT_FD10N_0100_OKCD_OCC3)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC8)

Call SelectCheckbox("NORMBL", 0, DT_FD10N_1000_STANDARD_DOCUMENTS, False)
Call SetTextbox("Fiscal year","RGJAHR-LOW","",DT_FD10N_1000_FISCAL_YEAR_OCC1,False)
Call SetTextbox("Document number","RBELNR-LOW","",DT_FD10N_1000_DOCUMENT_NUMBER,False)
Call SetTextbox("Company code","RBUKRS-LOW","",DT_FD10N_1000_COMPANY_CODE_OCC3,False)
Call SetTextbox("Document type","RBLART-LOW","",DT_FD10N_1000_DOCUMENT_TYPE,False)
Call SetTextbox("Correspondence","REVENT","",DT_FD10N_1000_CORRESPONDENCE,False)
Call ClickButton("Execute   \(F8\)", False)
Call PressEnter() 
Call ClickButtonIfExist("Yes", True)

Call SetTextbox("Output Device","USR01-SPLD","",DT_FD10N_1100_OUTPUT_DEVICE,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)", True)
Call TakeScreenShot
Call ClickButtonIfExist("Yes", True)
Call TakeScreenShot

' SetTextboxPopupIfExist(textboxName, attachedText, strTextboxValue)
Call SetTextboxPopupIfExist("USR01-SPLD","Output Device",DT_FD10N_1100_OUTPUT_DEVICE)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)", True)
Call ClickButtonIfExist("Yes", True)

Call ClickButton("Back   \(F3\)", False)
Call ClickButton("Back   \(F3\)", False)

''''--------TransactionCode-FB10N----------''''

Call SetTcode(DT_FD10N_0100_OKCD_OCC5)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC10)

Call SetTextbox("Customer","SO_KUNNR-LOW","",DT_FD10N_1000_CUSTOMER_OCC1,False)
Call SetTextbox("Company code","SO_BUKRS-LOW","",DT_FD10N_1000_COMPANY_CODE_OCC4,False)
Call SetTextbox("Fiscal year","GP_GJAHR","",DT_FD10N_1000_FISCAL_YEAR_OCC2,False)
Call PressEnter()
Call ClickButton("Execute   \(F8\)", False)
'
''Call SelectRowGuiGrid("", 0, "Period", DT_FD10N_ROW, False)
''Call TakeScreenShot
''
'' DoubleClickGuiGridCell(gridTitle, gridIndex, rowNumber, columnName, blnIsItPopup)
Call DoubleClickGuiGridCell("",0,DT_FD10N_ROW,"Period",False)
Call TakeScreenShot()

''Call ClickButtonIfExist("btn\[2\]", False)
'Call ClickButtonIfExist("Call up line item report   \(F2\)", False)

'Call ClickButton("btn\[38\]", False)
Call ClickButton("Set filter   \(Ctrl\+Shift\+F2\)", False)
Call SelectRowGuiTable("SAPLSKBHTC_FIELD_LIST_820", "Content", "Document Number", True)
Call ClickButton("APP_WL_SING", True)
'Call ClickButton("btn\[0\]", True)
Call ClickButton("Copy   \(Enter\)", True)
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_DOC_NO,True)
Call ClickButton("Execute   \(Enter\)", True)
Call VerifyifGuiLabelExists_ByIndex(DT_FD10N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_17_BALANCE_OCC1,0)
Call VerifyifGuiLabelExists_ByIndex(DT_DOC_NO,0)


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




