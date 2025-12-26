		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.07.01.03.02 Post GL Adjustment (Both Ledgers)
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


gstrTestCaseName = "Test_09.07.01.03.02 Post GL Adjustment (Both Ledgers)"
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
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_0100_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()

Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_0302_PSTKY,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F02_0302_TEXT,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_0302_AMOUNT,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_0302_ACCOUNT,False)
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_0302_AMOUNT_OCC1,False)
Call SetTextboxNoLabel("BSEG-ZTERM","",DT_F02_0302_PAYT_TERMS,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F02_0302_TEXT_OCC1,False)

Call PressEnter()
Call TakeScreenShot
Call SelectMenuBar("Document;Simulate")
Call TakeScreenShot

Call VerifyTextBoxContent("C","RF05A-AZSAL", 0, DT_F02_0700_CHECK_TEXT_OF_C, False)

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_DOC_1_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_1_OUTPUT&" was posted in company code GR02")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_1_OUTPUT",DT_DOC_1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)


Call SetTextbox("Period","BKPF-MONAT","",ConvertDoubleDigit(Cstr(MOnth(DT_F02_0100_PERIOD_OCC1))),False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_F02_0100_POSTING_DATE_OCC1),False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F02_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("Type","BKPF-BLART","",DT_F02_0100_TYPE_OCC1,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F02_0100_DOCUMENT_DATE_OCC1),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F02_0100_CURRENCYRATE_OCC1,False)

Call SetTextbox("Period","BKPF-MONAT","",ConvertDoubleDigit(Cstr(MOnth(DT_F02_0100_PERIOD_OCC2))),False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_F02_0100_POSTING_DATE_OCC1),False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F02_0100_COMPANY_CODE_OCC2,False)
Call SetTextbox("Type","BKPF-BLART","",DT_F02_0100_TYPE_OCC1,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F02_0100_DOCUMENT_DATE_OCC2),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F02_0100_CURRENCYRATE_OCC2,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F02_0100_REFERENCE_OCC1,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_0100_PSTKY_OCC1,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_0100_ACCOUNT_OCC1,False)
Call TakeScreenShot
Call PressEnter()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_0300_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F02_0300_TAX_CODE,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F02_0300_TEXT,False)

Call SetTextbox("Business Area","COBL-GSBER","",DT_F02_1006_BUSINESS_AREA,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_F02_1006_COST_CENTER,False)
Call SetTextbox("Order","COBL-AUFNR","",DT_F02_1006_ORDER,False)
Call SetTextbox("Profit Center","COBL-PRCTR","",DT_F02_1006_PROFIT_CENTER,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_0300_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_0300_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()

Call ClickButtonifExist("Continue   \(Enter\)",True)
Wait 1
Call ClickButtonifExist("Continue   \(Enter\)",True)


Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_0300_AMOUNT_OCC1,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F02_0300_TAX_CODE_OCC1,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F02_0300_TEXT_OCC1,False)

Call SetTextbox("Business Area","COBL-GSBER","",DT_F02_1006_BUSINESS_AREA_OCC1,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_F02_1006_COST_CENTER_OCC1,False)
Call SetTextbox("Order","COBL-AUFNR","",DT_F02_1006_ORDER_OCC1,False)
Call SetTextbox("Profit Center","COBL-PRCTR","",DT_F02_1006_PROFIT_CENTER_OCC1,False)
Call PressEnter()
Call TakeScreenShot

Call ClickButtonifExist("Continue   \(Enter\)",True)
Wait 1
Call ClickButtonifExist("Continue   \(Enter\)",True)

Call SelectMenuBar("Document;Simulate")
Call TakeScreenShot

Call VerifyTextBoxContent("C","RF05A-AZSAL", 0, DT_F02_0700_CHECK_TEXT_OF_C_OCC1, False)

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_DOC_2_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_2_OUTPUT&" was posted in company code GR02")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_2_OUTPUT",DT_DOC_2)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call LogOff()
Call FinalStatus()

