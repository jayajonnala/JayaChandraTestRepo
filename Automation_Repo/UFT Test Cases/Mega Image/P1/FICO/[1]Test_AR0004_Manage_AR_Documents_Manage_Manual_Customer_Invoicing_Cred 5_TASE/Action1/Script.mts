
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AR0004_Manage_AR_Documents_Manage_Manual_Customer_Invoicing_Cred 5
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


gstrTestCaseName = "Test_AR0004_Manage_AR_Documents_Manage_Manual_Customer_Invoicing_Cred 5"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\FICO\TASE_DT_AR0004_Manage_AR_Documents_Manage_Manual_Customer_Invoicing_Cred.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'''''--------TransactionCode-FB70----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

''Call SetTextbox("Company Code","BKPF-BUKRS","",DT_COMP_CODE,True)
Call SetTextboxPopupIfExist("BKPF-BUKRS", "Company Code", DT_COMP_CODE)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)", True)

Call SetTextbox("Customer","INVFO-ACCNT","",DT_FB70_0510_CUSTOMER,False)
Call SetTextbox("Invoice date","INVFO-BLDAT","",ConvertDate(DT_FB70_0510_INVOICE_DATE),False)
Call SetTextbox("Reference","INVFO-XBLNR","",DT_FB70_0510_REFERENCE,False)
Call SetTextbox("Posting Date","INVFO-BUDAT","",ConvertDate(DT_FB70_0510_POSTING_DATE),False)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB70_0510_AMOUNT,False)
Call ClickButton("Switch Company Code   \(F7\)",False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_COMP_CODE,True)
Call TakeScreenShot
Call PressEnter()
Call PressEnter()
Call SetTextbox("HeaderText","INVFO-BKTXT","",DT_FB70_0550_HEADERTEXT,False)
Call TakeScreenShot
Call PressEnter()
Call SelectTab("TS", "Basic data", False)
Call SelectCheckbox("INVFO-XMWST", 0, DT_FB70_0510_CALCULATE_TAX, False)
Call SetTextbox("Text","INVFO-SGTXT","",DT_FD10N_0510_TEXT,False)
Call TakeScreenShot
Call SetTableData("SAPLFSKBTABLE", "G/L acct", 1, "", "", DT_FB70_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", 1, "", "", DT_FB70_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "Tax code", 1, "", "", DT_FB70_0100_TABLECELL_TAX_CODE_0, False)
Call SetTableData("SAPLFSKBTABLE", "Text", 1, "", "", DT_FB70_0100_TABLECELL_TEXT_0_OCC2, False)
Call SetTableData("SAPLFSKBTABLE", "Business area", 1, "", "", DT_FB70_0100_TABLECELL_BUSINESS_AREA_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", 1, "", "", DT_FB70_0100_TABLECELL_COST_CENTER_0, False)
Call TakeScreenShot
Call PressEnter()
Call PressEnter()
Call TakeScreenShot
Call SelectTab("TS", "Payment", False)
Call GetTextboxValue("INVFO-ZBD1T", "", "DT_GET_TEXT_OF_DAYS_OUTPUT", False)
Call TakeScreenShot
Call ClickButton("Simulate Document Posting   \(F9\)", False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)", False)
Call GetStatusBar("item1", "DT_DOC_NO_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_NO_OUTPUT&" was posted in company code RO02")
Call ClickButtonIfExist("Cancel   \(F12\)", True)

'''''''--------TransactionCode-/FB03----------''''

Call SetTcode(DT_FB70_1200_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call SetTextbox("Document Number","RF05L-BELNR","",DT_DOC_NO_OUTPUT,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB70_100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FB70_100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()
Call GetTextboxValue("BKPF-XBLNR", "", "DT_GET_TEXT_OF_REFERENCE_OUTPUT", False)
Call VerifyGridCellContent("", 1, "ZUONR", 0, DT_GET_TEXT_OF_REFERENCE_OUTPUT)
Call VerifyGridCellContent("", 1, "Account", 0, DT_FB70_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "Account", 0, DT_FB70_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 1, "TAX CODE", 0, DT_FB70_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ)
Call VerifyGridCellContent("", 2, "TAX CODE", 0, DT_FB70_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MWSKZ)
Call TakeScreenShot


''''''''--------TransactionCode-/nfbl5n----------''''

Call SetTcode(DT_FB70_750_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)

Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_FB70_1000_CUSTOMER_ACCOUNT,False)
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_FB70_1000_COMPANY_CODE,False)
Call SetTextbox("Posting date","SO_BUDAT-LOW","",ConvertDate(DT_FB70_1000_POSTING_DATE),False)
Call TakeScreenShot
Call PressEnter()
Call ClickButton("Execute   \(F8\)", False)
Call TakeScreenShot
Call ClickButton("Set filter   \(Ctrl\+Shift\+F2\)", False)
Call ClickButton("Find", True)
Call SetTextbox("Find","GD_SEARCHSTR","",DT_FB70_850_FIND,True)
Call TakeScreenShot
Call PressEnter()
Call ClickButton("Show sel. fields \(CTRL\+F3\)",True)
Call ClickButton("Copy   \(Enter\)",True)
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_DOC_NO_OUTPUT,True)
Call ClickButton("Execute   \(Enter\)", True)
Call TakeScreenShot
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",4)
Call ClickButton("Change layout...   \(Ctrl\+F8\)", False)
Call SelectRowGuiTableByRow("SAPLSKBHTC_FIELD_LIST",1, True)
Call ClickButton("Show Sel. Fields \(Ctrl\+F3\)", True)
Call SelectRowGuiTableByRow("SAPLSKBHTC_FIELD_LIST",1, True)
Call ClickButton("Show Sel. Fields \(Ctrl\+F3\)", True)
Call SelectRowGuiTableByRow("SAPLSKBHTC_FIELD_LIST",71, True)
Call ClickButton("Show Sel. Fields \(Ctrl\+F3\)", True)
Call SelectRowGuiTableByRow("SAPLSKBHTC_FIELD_LIST",17, True)
Call ClickButton("Show Sel. Fields \(Ctrl\+F3\)", True)
Call TakeScreenShot
Call ClickButton("Copy   \(Enter\)", True)
Call VerifyifGuiLabelExists_ByIndex(DT_GET_TEXT_OF_REFERENCE_OUTPUT,0)
Call VerifyifGuiLabelExists_ByIndex(DT_DOC_NO_OUTPUT,0)
'Call VerifyifGuiLabelExists_ByIndex(lcase(DT_FB70_120_CHECK_TEXT_OF_TEXT),0)
Call VerifyifGuiLabelExists_ByIndex(ConvertDate(DT_CHECK_BASELINE_DATE),0)
'Updated Due date script for the layout change
Call SetHorizontalScrollBar (50, False)
Call VerifyifGuiLabelExists_ByIndex(ConvertDate(DT_FB70_120_CHECK_TEXT_OF_DUE_DATE),0)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)", False)
Call ClickButton("Back   \(F3\)", False)

''''''''--------TransactionCode-/FB12----------''''

Call SetTcode(DT_FB70_120_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)

Call SetTextbox("Company Code","RF022-BUKRS","",DT_FB70_100_COMPANY_CODE_OCC2,False)
Call PressEnter()
Call ClickButtonIfExist("Restrict Values   \(Shift\+F5\)", True)
Call SetTextbox("Correspondence","G_SELFLD_TAB-LOW","",DT_FB70_0220_CORRESPONDENCE,True)
Call TakeScreenShot
Call PressEnter()
Call PressEnter()
Call SetTextbox("Document Number","RF022-BELNR","",DT_DOC_NO_OUTPUT,True)
Call SetTextbox("Fiscal Year","RF022-GJAHR","",DT_FB70_1001_FISCAL_YEAR,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)", True)
Call GetStatusBar("MessageType", "MESSAGETYPE_S_OUTPUT")
Call VerifyStatusBar(DT_FB70_100_CHECK_TEXT_OF_STATUSBAR)

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




