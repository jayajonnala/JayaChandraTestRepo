		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_PRE_1_09.12.02.03.03 Manage Account Receivable Monthly Closing
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
	GetRowNo= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrTestCaseName = "Test_PRE_1_09.12.02.03.03 Manage Account Receivable Monthly Closing"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'

'''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''--------TransactionCode-FD32----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SelectCheckbox("RF02L-D0105", 0, DT_FD32_0100_OVERVIEW, False)
Call SelectCheckbox("RF02L-D0110", 0, DT_FD32_0100_ADDRESS, False)
Call SelectCheckbox("RF02L-D0120", 0, DT_FD32_0100_CENTRAL_DATA, False)
Call SelectCheckbox("RF02L-D0210", 0, DT_FD32_0100_STATUS, False)
Call SelectCheckbox("RF02L-D0220", 0, DT_FD32_0100_PAYMENT_HISTORY, False)

Call  SetTextbox("Customer","RF02L-KUNNR", 0, DT_FD32_0100_CUSTOMER, False)
Call  SetTextbox("Credit control area","RF02L-KKBER", 0, DT_FD32_0100_CREDIT_CONTROL_AREA, False)

Call ClickButton("Select All   \(F7\)",False)
Call ClickButton("Next screen   \(Shift\+F1\)",False)
Call ClickButton("Next screen   \(Shift\+F1\)",False)
Call ClickButton("Next screen   \(Shift\+F1\)",False)
Call SetTextbox("Risk category","KNKK-CTLPC", 0, DT_FD32_0210_RISK_CATEGORY, False)
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call PressEnter()
'Wait(3)
Call ClickButtonIfExist("Continue   \(Enter\)",False)
'
Call ClickButton("Select All   \(F7\)",False)
Call ClickButton("Next screen   \(Shift\+F1\)",False)
Call ClickButton("Next screen   \(Shift\+F1\)",False)
Call ClickButton("Next screen   \(Shift\+F1\)",False)
Call  SetTextbox("Risk category","KNKK-CTLPC", 0, DT_FD32_0210_RISK_CATEGORY_OCC1, False)
Call ClickButton("Save   \(Ctrl\+S\)",False)
'Wait(3)
Call PressEnter()
Call ClickButtonIfExist("Continue   \(Enter\)",False)

'''''''--------TransactionCode-FB70----------''''

Call SetTcode(DT_FD32_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FD32_1000_COMPANY_CODE,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot

'Call SetTextbox("Invoice date","INVFO-BLDAT", "", "04.07.2022", False)
'Call SetTextbox("Posting Date","INVFO-BUDAT", "", "04.07.2022", False)
Call SetTextbox("Invoice date","INVFO-BLDAT", "", ConvertDate(DT_FD32_0510_INVOICE_DATE), False)
Call SetTextbox("Posting Date","INVFO-BUDAT", "", ConvertDate(DT_FD32_0510_POSTING_DATE), False)
Call SetTextbox("Customer","INVFO-ACCNT","",DT_FD32_0510_CUSTOMER,False)
Call SetTextbox("Amount","INVFO-WRBTR", "", DT_FD32_0510_AMOUNT, False)
Call SetTextbox("Text","INVFO-SGTXT", "", DT_FD32_0510_TEXT, False)
Call TakeScreenShot
Call SetComboByKey("INVFO-MWSKZ",DT_FD32_0510_TAX_AMOUNT)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call SelectCheckbox("INVFO-XMWST","0",DT_FD32_0510_CALCULATE_TAX,False)
Call TakeScreenShot

Call SetTableData("SAPLFSKBTABLE", "G/L acct", "1", "", "", DT_FD32_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "1", "", "", DT_FD32_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "Text", "1", "", "", DT_FD32_0100_TABLECELL_TEXT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Business area", "1", "", "", DT_FD32_0100_TABLECELL_BUSINESS_AREA_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", "1", "", "", DT_FD32_0100_TABLECELL_COST_CENTER_0, False)
Call PressEnter()
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)

Call GetStatusBar("item1","DT_FD32_1000_CHECK_TEXT_OF_COMPANY_CODE_OUTPUT")
Call VerifyStatusBar("Document "&DT_FD32_1000_CHECK_TEXT_OF_COMPANY_CODE_OUTPUT&" was posted in company code GR02")
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet("DT_FD32_1000_CHECK_TEXT_OF_COMPANY_CODE_OUTPUT",DT_FD32_1000_CHECK_TEXT_OF_COMPANY_CODE)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FD32_1000_COMPANY_CODE_OCC1,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SetTextbox("Invoice date","INVFO-BLDAT", "", ConvertDate(DT_FD32_0510_INVOICE_DATE_OCC1), False)
Call SetTextbox("Posting Date","INVFO-BUDAT", "", ConvertDate(DT_FD32_0510_POSTING_DATE_OCC1), False)
Call SetTextbox("Customer","INVFO-ACCNT","",DT_FD32_0510_CUSTOMER_OCC1,False)
Call SetTextbox("Amount","INVFO-WRBTR", "", DT_FD32_0510_AMOUNT_OCC1, False)
Call SetTextbox("Text","INVFO-SGTXT", "", DT_FD32_0510_TEXT_OCC1, False)
Call TakeScreenShot
Call SetComboByKey("INVFO-MWSKZ",DT_FD32_0510_TAX_AMOUNT_OCC1)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call SelectCheckbox("INVFO-XMWST","0",DT_FD32_0510_CALCULATE_TAX,False)
Call TakeScreenShot

Call SetTableData("SAPLFSKBTABLE", "G/L acct", "1", "", "", DT_FD32_0100_TABLECELL_GL_ACCT_0_OCC1, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "1", "", "", DT_FD32_0100_TABLECELL_AMOUNT_IN_DOCCURR_0_OCC1, False)
Call SetTableData("SAPLFSKBTABLE", "Text", "1", "", "", DT_FD32_0100_TABLECELL_TEXT_0_OCC1, False)
Call SetTableData("SAPLFSKBTABLE", "Business area", "1", "", "", DT_FD32_0100_TABLECELL_BUSINESS_AREA_0_OCC1, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", "1", "", "", DT_FD32_0100_TABLECELL_COST_CENTER_0_OCC1, False)
Call PressEnter()
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call SetComboByKey("INVFO-BLART",DT_FD32_0510_DOCUMENT_TYPE)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)

Call GetStatusBar("item1","DT_FD32_1000_CHECK_TEXT_OF_COMPANY_CODE_OCC2_OUTPUT")
Call VerifyStatusBar("Document "&DT_FD32_1000_CHECK_TEXT_OF_COMPANY_CODE_OCC2_OUTPUT&" was posted in company code GR02")
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet("DT_FD32_1000_CHECK_TEXT_OF_COMPANY_CODE_OCC2_OUTPUT",DT_FD32_1000_CHECK_TEXT_OF_COMPANY_CODE_OCC2)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call ClickButton("Cancel   \(F12\)",True)

''''''--------TransactionCode-ZFIAR_F103----------''''

Call SetTcode(DT_FD32_0100_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Customer account","DD_KUNNR-LOW", "", DT_FD32_1000_CUSTOMER_ACCOUNT, False)
Call SetTextbox("Company code","DD_BUKRS-LOW", "", DT_FD32_1000_COMPANY_CODE_OCC2, False)
Call SetTextbox("Provision method","RSMET", "", DT_FD32_1000_PROVISION_METHOD, False)
Call SetTextbox("Spec\.G/L indic\.for dbtfl\. rec\.","B-UMSKZ", "", DT_FD32_1000_SPECGL_INDICFOR_DBTFL_REC, False)
Call SetTextbox("Posting date","B-BUDAT", "", ConvertDate(DT_FD32_1000_POSTING_DATE), False)
Call SetTextbox("Document date","B-BLDAT", "", ConvertDate(DT_FD32_1000_DOCUMENT_DATE), False)

Call ClickButton("%_BELNR_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_FD32_0120_CHECK_TEXT_OF_NO_NAME,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_FD32_0120_CHECK_TEXT_OF_NO_NAME_OCC1,True)
Call ClickButton("Copy   \(F8\)",False)
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call SelectCheckbox("P_FILKD","0",DT_FD32_1000_POST_TO_BRANCH,False)
Call TakeScreenShot
Call SelectCheckbox("B-INPUT","0",DT_FD32_1000_CREATE_BATCH_INPUT_SESSION,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call VerifyifGuiLabelExistsByRelativeid(DT_FD32_0120_CHECK_TEXT_OF_NO_NAME, "wnd\[0\]/usr/lbl\[42,12\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_FD32_0120_CHECK_TEXT_OF_NO_NAME_OCC1, "wnd\[0\]/usr/lbl\[42,13\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_FD32_0120_CHECK_TEXT_OF_EUR, "wnd\[0\]/usr/lbl\[57,12\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_FD32_0120_CHECK_TEXT_OF_EUR_OCC1, "wnd\[0\]/usr/lbl\[57,13\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_FD32_0120_CHECK_TEXT_OF_NO_NAME_OCC2, "wnd\[0\]/usr/lbl\[88,18\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_FD32_0120_CHECK_TEXT_OF_NO_NAME_OCC3, "wnd\[0\]/usr/lbl\[88,12\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_FD32_0120_CHECK_TEXT_OF_NO_NAME_OCC4, "wnd\[0\]/usr/lbl\[88,13\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_FD32_0120_CHECK_TEXT_OF_C, "wnd\[0\]/usr/lbl\[112,9\]")

Call Logoff()
Call FinalStatus ()



