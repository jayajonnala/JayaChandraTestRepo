		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.04.05.01.03 Issuing Manual Direct Vendor Credit Note 
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

gstrTestCaseName = "Test_09.04.05.01.03 Issuing Manual Direct Vendor Credit Note"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'

''''''--------------login----------------'''''
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''''--------TransactionCode-FB60----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB60_1000_COMPANY_CODE,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SetTextbox("Invoice date", "INVFO-BLDAT", "", ConvertDate(DT_FB60_0010_INVOICE_DATE), False)
Call SetTextboxNoLabel("INVFO-ACCNT","",DT_FB60_0010_VENDOR,False)
Call SetComboByKey("INVFO-BLART",DT_FB60_0010_DOCUMENT_TYPE)
Call TakeScreenShot
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB60_0010_AMOUNT,False)
Call SelectCheckbox("INVFO-XMWST","0",DT_FB60_0010_CALCULATE_TAX,False)
Call TakeScreenShot
Call SetComboByKey("INVFO-MWSKZ",DT_FB60_0010_INVFOMWSKZ)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenSHot
Call SetTextbox("Text","INVFO-SGTXT","",DT_FB60_0010_TEXT,False)

Call SetTableData("SAPLFSKBTABLE", "G/L acct", "1", "", "", DT_FB60_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Text", "1", "", "", DT_FB60_0100_TABLECELL_TEXT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "1", "", "", DT_FB60_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE","Business area", "1", "", "", DT_FB60_0100_TABLECELL_BUSINESS_AREA_0, False)
Call SetTableData("SAPLFSKBTABLE","Cost center", "1", "", "", DT_FB60_0100_TABLECELL_COST_CENTER_0, False)
Call PressEnter()
Call TakeScreenShot
Call ClickCellTableByRowNo("SAPLFSKBTABLE", "Long Txt", "1", False)

Call ClickButton("Other languages",True)
Call SetTextbox("Language","SVALD-VALUE","",DT_FB60_0300_LANGUAGE,True)
Call PressEnter()
CAll SetSpecialTextbox("1st line","EENO_DYNP-ZEILE", "0", DT_FB60_1001_1ST_LINE, True)
Call ClickButton("Copy text   \(F5\)",True)
Call ClickButton("Simulate Document Posting   \(F9\)",False)
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)",False)

Call GetStatusBar("item1","DT_DOC_NUMBER_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_NUMBER_OUTPUT&" was posted in company code GR02")
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_NUMBER_OUTPUT",DT_DOC_NUMBER)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call ClickButtonIfExist("Cancel   \(F12\)",False)


'''--------TransactionCode-F.62 ----------''''

Call SetTcode(DT_FB60_1100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SelectCheckbox("NORMBL", 0, DT_FB60_1000_STANDARD_DOCUMENTS, False)
Call SetTextbox("Fiscal year","RGJAHR-LOW","",Year(DT_FB60_1000_FISCAL_YEAR),False)
Call SetTextbox("Company code","RBUKRS-LOW","",DT_FB60_1000_COMPANY_CODE_OCC2,False)
Call SetTextbox("Document number","RBELNR-LOW","",DT_FB60_1000_DOCUMENT_NUMBER,False)
Call SetTextbox("Document type","RBLART-LOW","",DT_FB60_1000_DOCUMENT_TYPE,False)
Call SetTextbox("Correspondence","REVENT","",DT_FB60_1000_CORRESPONDENCE,False)
Call SetTextbox("Delete if finished since","RDELDAYS","","",False)
'Call SetTextbox("Delete if finished since","RDELDAYS","",DT_FB60_1000_DELETE_IF_FINISHED_SINCE,False)
Call SetTextbox("Log to printer","PRDEST","",DT_FB60_1000_LOG_TO_PRINTER,False)
Call ClickBUtton("Execute   \(F8\)",False)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call ClickButtonIfExist("Yes",True)


Call SetTextbox("Output Device","USR01-SPLD","",DT_FB60_1000_OUTPUT_TO_PRINTER,False)
Call PressEnter()
Call TakeScreenShot

''''--------TransactionCode-SP02 ----------''''

Call SetTcode(DT_FB60_0120_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

'DT_FB75_0120_CHECK_TEXT_OF_NO_NAME_IP = DT_FB75_0120_CHECK_TEXT_OF_NO_NAME&"/"&Year(DT_FORMULA_YEAR)
Call SetFocusGuiLabel(DT_FB60_0120_CHECK_TEXT_OF_NO_NAME_OCC1,"347","56", False)
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenshot
Call VerifyifGuiLabelExistsByRelativeid(DT_FB60_0120_CHECK_TEXT_OF_NO_NAME_OCC1,"wnd\[0\]/usr/lbl\[49,3\]")
Call VerifyifGuiLabelExistsByRelativeid(ConvertDate(DT_FB60_0120_CHECK_TEXT_OF_NO_NAME),"wnd\[0\]/usr/lbl\[19,3\]")
Call TakeScreenshot
Call ClickLabel("X__PDF", "0", False)
Call TakeScreenshot

Call LogOff'
Call FinalStatus()
