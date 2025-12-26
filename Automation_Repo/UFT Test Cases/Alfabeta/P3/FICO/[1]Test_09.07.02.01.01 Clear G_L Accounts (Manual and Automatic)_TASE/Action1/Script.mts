		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.07.02.01.01 Clear G_L Accounts (Manual and Automatic)
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

gstrTestCaseName = "Test_09.07.02.01.01 Clear G_L Accounts (Manual and Automatic)"
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
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_0100_PSTKY,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F02_0100_DOCUMENT_DATE),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F02_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F02_0100_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F02_0100_DOCHEADER_TEXT,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_0100_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()


Call SetTextbox("Text","BSEG-SGTXT","",DT_F02_0300_TEXT,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_0300_AMOUNT,False)
Call ClickButton("Display Additional Data for Document Item   \(F7\)",False)
Call TakeScreenShot
Call SetTextbox("House Bank","BSEG-HBKID","",DT_F02_0330_HOUSE_BANK,False)
Call SetTextbox("/","BSEG-HKTID","",DT_F02_0330_BSEGHKTID,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_0330_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_0330_ACCOUNT,False)
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Text","BSEG-SGTXT","",DT_F02_0300_TEXT_OCC1,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_0300_AMOUNT_OCC1,False)
Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
Call TakeScreenShot
Call VerifyTextBoxContent("C", "RF05A-AZSAL", 0, DT_F02_0700_CHECK_TEXT_OF_C, False)


Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_DOC_1_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_1_OUTPUT&" was posted in company code GR02")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_1_OUTPUT",DT_DOC_1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

'''''''''--------TransactionCode-F-02----------''''
'''
Call SetTcode(DT_F02_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_F02_0100_POSTING_DATE_OCC1),False)
Call SetTextbox("Type","BKPF-BLART","",DT_F02_0100_TYPE_OCC1,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F02_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F02_0100_DOCUMENT_DATE_OCC1),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F02_0100_CURRENCYRATE_OCC1,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F02_0100_REFERENCE_OCC1,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F02_0100_DOCHEADER_TEXT_OCC1,False)
Call TakeScreenShot
Call CliCkButton("Post with Reference Document   \(Shift\+F9\)",False)

Call SelectCheckbox("RF05A-CPSTO", 0, DT_F02_0104_GENERATE_REVERSE_POSTING,False)
Call SetTextbox("Document Number","BKPF-BELNR","",DT_F02_0104_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F02_0104_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","BKPF-GJAHR","",Year(DT_F02_0104_FISCAL_YEAR),False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call VerifyTextBoxContent("C", "RF05A-AZSAL", 0, DT_F02_0700_CHECK_TEXT_OF_C_OCC1, False)

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call GetTextboxValue("MESSTXT1", 0, "DT_OUTPUT_DOC_OUTPUT", blnIsItPopup)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_OUTPUT_DOC_OUTPUT",DT_OUTPUT_DOC)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call PressEnter()
'
''''''''''--------TransactionCode-F-03----------''''
Call SetTcode(DT_F02_0100_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Clearing Date","BKPF-BUDAT","",ConvertDate(DT_F02_0131_CLEARING_DATE),False)
Call SetTextbox("Account","RF05A-AGKON","",DT_F02_0131_ACCOUNT,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F02_0131_COMPANY_CODE,False)
Call SetTextbox("Period","BKPF-MONAT","",ConvertDoubleDigit(Cstr(Month(DT_F02_0131_PERIOD))),False)
Call SetTextbox("Currency","BKPF-WAERS","",DT_F02_0131_CURRENCY,False)
Call SelectRadioButton("RF05A-XPOS1", "Document Number", False)

Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
Call SetTextbox("From","RF05A-SEL01","",DT_F02_0731_FROM,False)
Call SetTextbox("To","RF05A-SEL02","",DT_F02_0731_TO,False)

Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
Call VerifyTextBoxContent("Not assigned","RF05A-DIFFB", 0, DT_F02_6103_CHECK_TEXT_OF_NOT_ASSIGNED, False)

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_DOC_3_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_3_OUTPUT&" was posted in company code GR02")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_3_OUTPUT",DT_DOC_3)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

''''''''''--------TransactionCode-FAGLL03----------''''
Call SetTcode(DT_F02_0131_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_F02_1000_GL_ACCOUNT,False)
Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_F02_1000_COMPANY_CODE,False)
Call SelectRadioButton("X_AISEL","All items", False)


Call ClickButton("Custom Selections   \(Ctrl\+F1\)",false)
Call ActivateNodeGuiTree("0","#4;#1")
Call ClickButton("%_%%DYN001_%_APP_%-VALU_PUSH",Fasle)
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "", "", DT_F02_3010_TABLECELL_SINGLE_VALUE_0, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "", "", DT_F02_3010_TABLECELL_SINGLE_VALUE_1, True)
Call ClickButton("Copy   \(F8\)",True)

Call ClickButton("Save   \(Ctrl\+S\)",False)

Call ClickButtonIfExist("Execute   \(F8\)",False)
Call VerifyStatusBarMessageType("S")

Call VerifyGridCellContent("", 3, "DMSHB", 0, DT_F02_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_DMSHB)
Call LogOff()
Call FinalStatus()


'
