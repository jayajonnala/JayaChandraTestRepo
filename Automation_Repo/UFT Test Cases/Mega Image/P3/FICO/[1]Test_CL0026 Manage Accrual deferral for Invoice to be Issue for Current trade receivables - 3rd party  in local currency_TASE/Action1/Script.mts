'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_CL0026 Manage Accrual deferral for Invoice to be Issue for Current trade receivables - 3rd party in local currency  
'.................Author : TCS      
'................ Creation Date   
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


gstrTestCaseName = "Test_CL0026"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\FICO\TASE_DT_CL0026 Manage Accrual deferral for Invoice to be Issue for Current trade receivables - 3rd party  in local currency.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''----------------------Login----------------------------

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''--------TransactionCode-FBS1---------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FBS1_100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FBS1_100_ACCOUNT,False)
Call SetTextbox("SGL Ind","RF05A-NEWUM","",DT_FBS1_100_SGL_IND,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",DT_FBS1_100_DOCUMENT_DATE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_FBS1_100_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FBS1_100_COMPANY_CODE,False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",DT_FBS1_100_POSTING_DATE,False)
Call SetTextbox("Period","BKPF-MONAT","",DT_FBS1_100_PERIOD,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FBS1_100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_FBS1_100_REFERENCE,False)
Call SetTextbox("Doc.Header Text","BKPF-BKTXT","",DT_FBS1_100_DOCHEADER_TEXT,False)
Call SetTextbox("Reversal Reason","BKPF-STGRD","",DT_FBS1_100_REVERSAL_REASON,False)
Call SetTextbox("Reversal Date","BKPF-STODT","",DT_FBS1_100_REVERSAL_DATE,False)
Call TakeScreenShot
Call PressEnter()
If VerifyStatusBar("Period") Then
	Call TakeScreenShot
	Call PressEnter()
End If
Call PressEnter()
Call TakeScreenShot
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FBS1_303_PSTKY,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_FBS1_303_TEXT,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FBS1_303_ACCOUNT,False)
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_FBS1_303_ASSIGNMENT,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FBS1_303_AMOUNT,False)
Call SetTextbox("Due on","BSEG-ZFBDT","",DT_FBS1_303_DUE_ON,False)
Call TakeScreenShot
Call PressEnter()
Call PressEnter()
Call SetTextbox("Text","BSEG-SGTXT","",DT_FBS1_300_TEXT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_FBS1_300_TAX_CODE,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FBS1_300_AMOUNT,False)
Call TakeScreenShot
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_FBS1_1006_COST_CENTER,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_FBS1_1006_BUSINESS_AREA,False)
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)", False)
If VerifyStatusBar("Tax") Then
	Call TakeScreenShot
	Call PressEnter()
End If
Call GetStatusBar("item1", "DT_FBS1_DOC_OUTPUT")
Call TakeScreenShot

'''--------\FB08 --------------

Call SetTcode(DT_FBS1_100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Posting Date","BSIS-BUDAT","",DT_FBS1_105_POSTING_DATE,False)
Call SetTextbox("Document Number","RF05A-BELNS","",DT_FBS1_DOC_OUTPUT,False)
Call TakeScreenShot
Call PressEnter()
Call ClickButton("Post   \(Ctrl\+S\)", False)
Call GetStatusBar("item1", "DT_FB08_DOC_OUTPUT")
Call TakeScreenShot

''''--------\FB03 --------------

Call SetTcode(DT_FBS1_105_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Document Number","RF05L-BELNR","",DT_FBS1_DOC_OUTPUT,False)
''Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB70_0100_COMPANY_CODE,False)
''Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FB70_0100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()
Call SelectRowGuiGridbyRowNo("", 0, 1, False)
Call DoubleClickGuiGridCell("", 0, 1, "Account", False)
'''Call GetTextboxValue("BSEG-HKONT", 0, "DT_ZFIGL_UPLOAD_POST_303_GET_TEXT_OF_GL_ACC_OUTPUT_OUTPUT", False)
Call VerifyTextBoxContent("G/L Acc", "BSEG-HKONT", 0, DT_FBS1_303_CHECK_TEXT_OF_GL_ACC, False)
Call ClickButtonIfExist("Display Document Header   \(F5\)", False)
Call TakeScreenShot
Call VerifyTextBoxContent("TCode", "BKPF-TCODE", 0, DT_FBS1_1710_CHECK_TEXT_OF_TCODE, False)
Call VerifyTextBoxContent("Document Date", "BKPF-BLDAT", 0, DT_FBS1_1710_CHECK_TEXT_OF_DOCUMENT_DATE, False)
Call VerifyTextBoxContent("Posting Date", "BKPF-BUDAT", 0, DT_FBS1_1710_CHECK_TEXT_OF_POSTING_DATE, False)
Call VerifyTextBoxContent("Reversal Reason", "BKPF-STGRD", 0, DT_FBS1_1710_CHECK_TEXT_OF_REVERSAL_REASON, False)
Call ClickButtonIfExist("Continue/Confirm   \(Enter\)", True)

''''--------\FB03 --------------

Call SetTcode(DT_FBS1_105_OKCD)     
Call PressEnter()  
Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB08_DOC_OUTPUT,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("Display Document Header   \(F5\)", False)
Call TakeScreenShot
Call VerifyTextBoxContent("Document Type", "BKPF-BLART", 0, DT_FBS1_1710_CHECK_TEXT_OF_DOCUMENT_TYPE, False)
Call VerifyTextBoxContent("Document Date", "BKPF-BLDAT", 0, DT_FBS1_1710_CHECK_TEXT_OF_DOCUMENT_DATE_OCC2, False)
Call VerifyTextBoxContent("Posting Date", "BKPF-BUDAT", 0, DT_FBS1_1710_CHECK_TEXT_OF_POSTING_DATE_OCC2, False)
Call VerifyTextBoxContent("Reversal Reason", "BKPF-STGRD", 0, DT_FBS1_1710_CHECK_TEXT_OF_REVERSAL_REASON, False)
Call ClickButtonIfExist("Continue/Confirm   \(Enter\)", True)
Call SelectRowGuiGridbyRowNo("", 0, 1, False)
Call DoubleClickGuiGridCell("", 0, 1, "Account", False)
Call ClickButtonIfExist("Display Additional Data   \(F8\)", False)
Call TakeScreenShot
Call VerifyCheckBoxValue("BSEG-XNEGP", "ON")
Call ClickButtonIfExist("Cancel   \(F12\)", True)
Call ClickButtonIfExist("Display Next Item   \(Shift\+F7\)", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Account", "BSEG-HKONT", 0, DT_FBS1_300_CHECK_TEXT_OF_GL_ACCOUNT, False)
Call ClickButtonIfExist("Display Additional Data   \(F8\)", False)
Call TakeScreenShot
Call VerifyCheckBoxValue("BSEG-XNEGP", "ON")
Call ClickButtonIfExist("Continue/Confirm   \(Enter\)", True)

''''''-----//fbl5n screen--------------

Call SetTcode(DT_FBS1_100_OKCD_OCC3)     
Call PressEnter() 
Call TakeScreenShot
Call SelectRadioButton("X_AISEL", "All items", False)
Call SelectCheckbox("X_SHBV", 0, DT_FBS1_1000_SPECIAL_GL_TRANSACTIONS, False)
Call SelectCheckbox("X_NORM", 0, DT_FBS1_1000_NORMAL_ITEMS, False)
Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_FBS1_1000_CUSTOMER_ACCOUNT,False)
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_FBS1_1000_COMPANY_CODE,False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)", Falser)
Call ClickButtonIfExist("btn\[33\]", False)
Call PressEnter() 
Call ClickButtonIfExist("Exit   \(Shift\+F3\)", False)

''''''-----/faglb03 screen--------------

Call SetTcode(DT_FBS1_100_OKCD_OCC4)     
Call PressEnter() 
Call TakeScreenShot
Call SetTextbox("Account Number","RACCT-LOW","",DT_FBS1_1000_ACCOUNT_NUMBER,False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)", Falser)
Call TakeScreenShot
Call ClickButtonIfExist("Exit   \(Shift\+F3\)", False)

Call LogOff()
Call FinalStatus ()

