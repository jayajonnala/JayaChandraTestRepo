		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.07.02.03.01 Post GL With Clearing
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

gstrTestCaseName = "Test_09.07.02.03.01 Post GL With Clearing"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''''''--------TransactionCode-F-02----------''''
'''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_F02_0100_POSTING_DATE),False)
Call SetTextbox("Type","BKPF-BLART","",DT_F02_0100_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F02_0100_COMPANY_CODE,False)
Call SetTextbox("Period","BKPF-MONAT","",ConvertDoubleDigit(Cstr(MOnth(DT_F02_0100_PERIOD))),False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_0100_PSTKY,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F02_0100_DOCUMENT_DATE),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F02_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F02_0100_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F02_0100_DOCHEADER_TEXT,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_0100_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()


Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_0300_PSTKY,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F02_0300_TEXT,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_0300_AMOUNT,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_0300_ACCOUNT,False)
Call PressEnter()
Call TakeScreenShot


Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_0300_AMOUNT_OCC1,False)
Call PressEnter()
Call TakeScreenShot
Call SelectMenuBar("Document;Simulate")
Call TakeScreenShot
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_DOC_1_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_1_OUTPUT&" was posted in company code GR02")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_1_OUTPUT",DT_DOC_1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

''''''''--------TransactionCode-F-51----------''''
Call SetTcode(DT_F02_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F02_0122_POSTING_DATE),False)
Call SetTextbox("Type","BKPF-BLART","",DT_F02_0122_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F02_0122_COMPANY_CODE,False)
Call SetTextbox("Period","BKPF-MONAT","",ConvertDoubleDigit(Cstr(MOnth(DT_F02_0122_PERIOD))),False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_0122_PSTKY,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F02_0122_DOCUMENT_DATE),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F02_0122_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F02_0122_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F02_0122_DOCHEADER_TEXT,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_0122_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_0300_AMOUNT_OCC2,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F02_0300_TEXT_OCC1,False)
Call ClickButtonIfExist("Display Additional Data for Document Item   \(F7\)",False)
Call TakeScreenShot

Call SetTextbox("House Bank","BSEG-HBKID","",DT_F02_0330_HOUSE_BANK,False)
Call SetTextbox("/","BSEG-HKTID","",DT_F02_0330_BSEGHKTID,False)
Call TakeScreenShot

Call ClickButton("Choose open items   \(F6\)",False)

Call SelectRadioButton("RF05A-XPOS1","Document Number",False)
Call SetTextbox("Account Type","RF05A-AGKOA","",DT_F02_0710_ACCOUNT_TYPE,False)
Call SetTextbox("Account","RF05A-AGKON","",DT_F02_0710_ACCOUNT,False)

Call ClickButton("Process Open Items   \(Shift\+F4\)",False)

Call SetTextbox("From","RF05A-SEL01","",DT_F02_0731_FROM,False)

Call ClickButton("Process Open Items   \(Shift\+F4\)",False)

Call VerifyTableCellContent(1, "EUR Gross", "SAPDF05XTC_6103", DT_F02_6103_CHECK_TEXT_OF_TABLECELL______EUR_GROSS_0)
Call VerifyTableCellContent(1, "Document Number", "SAPDF05XTC_6103", DT_F02_6103_CHECK_TEXT_OF_TABLECELL_DOCUMENT_NUMBER_0)

Call SelectMenuBar("Document;Simulate")
Call TakeScreenShot
Call VerifyTextBoxContent("C","RF05A-AZSAL", 0, DT_F02_0700_CHECK_TEXT_OF_C_OCC1, False)

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_DOC_2_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_2_OUTPUT&" was posted in company code GR02")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_2_OUTPUT",DT_DOC_2)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

''''''''--------TransactionCode-FB03----------''''
Call SetTcode(DT_F02_0122_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Document Number","RF05L-BELNR","",DT_F02_0100_DOCUMENT_NUMBER	,False) 
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_F02_0100_COMPANY_CODE_OCC1,False) 
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",Year(DT_F02_0100_FISCAL_YEAR),False) 
Call PressEnter()     
Call TakeScreenShot

Call ClickButton("Display Document Header   \(F5\)",False)

Call VerifyTextBoxContent("Document type","BKPF-BLART", 0, DT_F02_1710_CHECK_TEXT_OF_DOCUMENT_TYPE, True)
Call VerifyTextBoxContent("Reference","BKPF-XBLNR", 0, DT_F02_1710_CHECK_TEXT_OF_REFERENCE, True)

Call ClickButton("Continue/Confirm   \(Enter\)",True)

Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

Call VerifyGridCellContent("", 1, "KTONR", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)

Call VerifyGridCellContent("", 1, "Text", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)

Call VerifyGridCellContent("", 1, "Amount", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)

Call VerifyGridCellContent("", 1, "Currency", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_RF05A_UBAZW)
Call VerifyGridCellContent("", 2, "Currency", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_RF05A_UBAZW)

Call Logoff()
Call FinalStatus()




