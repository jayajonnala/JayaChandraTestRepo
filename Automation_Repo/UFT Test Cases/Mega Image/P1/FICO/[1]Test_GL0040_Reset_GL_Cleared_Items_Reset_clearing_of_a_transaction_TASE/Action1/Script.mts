
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_GL0040_Reset_GL_Cleared_Items_Reset_clearing_of_a_transaction
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

gstrTestCaseName = "Test_GL0040_Reset_GL_Cleared_Items_Reset_clearing_of_a_transaction"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''''''''''--------TransactionCode-F02----------''''
''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_100_ACCOUNT,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F02_100_DOCUMENT_DATE),False)
Call SetTextbox("Type","BKPF-BLART","",DT_F02_100_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F02_100_COMPANY_CODE,False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_F02_100_POSTING_DATE),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F02_100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F02_100_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F02_100_DOCHEADER_TEXT,False)

Call TakeScreenShot
Call FocusTextBox("Account","RF05A-NEWKO",False)
Call PressEnter()


Call SetTextbox("Text","BSEG-SGTXT","",DT_F02_300_TEXT,False)
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_F02_300_ASSIGNMENT,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_300_AMOUNT,False)
Call TakeScreenShot
'Call SetTextbox("Profit Center","COBL-PRCTR","",DT_F02_1006_PROFIT_CENTER,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F02_1006_TAX_CODE,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_F02_1006_BUSINESS_AREA,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_F02_1006_COST_CENTER,False)
Call TakeScreenShot
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_300_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_300_ACCOUNT,False)
Call TakeScreenShot
' FocusTextBox(attachedText, textboxName, blnIsItPopup)
Call FocusTextBox("Account","RF05A-NEWKO",False)
Call PressEnter()
'Call PressEnter()

Call SetTextbox("Text","BSEG-SGTXT","",DT_F02_300_TEXT_OCC2,False)
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_F02_300_ASSIGNMENT_OCC2,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_300_AMOUNT_OCC2,False)
Call TakeScreenShot
Call FocusTextBox("Text","BSEG-SGTXT",False)
Call PressEnter()

' ClickButtonIfExist(tooltipOrButtonName, blnIsItPopup)
Call ClickButtonIfExist("Display Document Overview   \(Shift\+F2\)",False)
Call PressEnter()
Call TakeScreenShot

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_DOC_NO1_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_NO1_OUTPUT&" was posted in company code RO02")

Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_NO1_OUTPUT",DT_DOC_NO1_OUTPUT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call TakeScreenShot
Call ClickButtonIfExist("Cancel   \(F12\)",True)


'''''''''--------TransactionCode-FBR2----------''''
'''
Call SetTcode(DT_EXPECTEDTRANSACTIONCODE_OCC2)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

CAll SelectCheckbox("RF05A-CPSTO", 0, DT_F02_104_GENERATE_REVERSE_POSTING, False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call PressEnter()

Call ClickButtonIfExist("Display Document Overview   \(Shift\+F2\)",False)
Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("Display Document Overview   \(Shift\+F2\)",False)
Call PressEnter()
Call TakeScreenShot

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetStatusBar("item1","DT_DOC_NO2_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_NO2_OUTPUT&" was posted in company code RO02")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_NO2_OUTPUT",DT_DOC_NO2_OUTPUT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

''''''''--------TransactionCode-F-03----------''''
''
Call SetTcode(DT_F02_100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Account","RF05A-AGKON","",DT_F02_131_ACCOUNT,False)
Call SetTextbox("Clearing Date","BKPF-BUDAT","",ConvertDate(DT_F02_131_CLEARING_DATE),False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F02_131_COMPANY_CODE,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call ClickButtonIfExist("Select All",False)
Call ClickButtonIfExist("Deactivate Items",False)
Call TakeScreenShot()
Call ClickButtonIfExist("Field content search",False)

Call SelectRadioButton("RF05A-XPOS1","Document Number", True)
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)

Call SetTextbox("From","RF05A-SEL01",0,DT_F02_0731_FROM,True)
Call SetTextbox("From","RF05A-SEL01",1,DT_F02_0731_FROM_OCC2,True)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call ClickButtonIfExist("Select All",False)
Call ClickButtonIfExist("Activate Items",False)
Call TakeScreenShot()

Call ClickButtonIfExist("Display Document Overview   \(Shift\+F2\)",False)
Call TakeScreenShot
Call SelectMenuBar("Document;Simulate")
Call TakeScreenShot
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot

Call TakeScreenShot
Call GetStatusBar("item1","DT_CL_DOC_NO_OUTPUT")
Call VerifyStatusBar("Document "&DT_CL_DOC_NO_OUTPUT&" was posted in company code RO02")

Call WriteRunTimeDataToExcelGlobalSheet ("DT_CL_DOC_NO_OUTPUT",DT_CL_DOC_NO_OUTPUT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'''''''''--------TransactionCode-FAGLL03----------''''
'''
Call SetTcode(DT_F02_131_OKCD)     
Call PressEnter()     
Call TakeScreenShot


Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_F02_1000_GL_ACCOUNT,False)

Call SelectRadioButton("X_AISEL","All Items", False)
Call TakeScreenShot()

Call SetTextbox("Posting Date","SO_BUDAT-LOW","",ConvertDAte(DT_F02_1000_POSTING_DATE),False)
Call SetTextbox("to","SO_BUDAT-HIGH","",ConvertDAte(DT_F02_1000_TO),False)

'Call SendKey("{F8}")
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenSHot()

CAll ClickLabel("Clrng doc\.", 0, FAlse)
Call TAkeScreenShot()
Call ClickBUtton("Set Filter   \(Ctrl\+Shift\+F2\)",FAlse)

Call SetTextbox("Clearing Document","%%DYN001-LOW","",DT_F02_100_CLEARING_DOCUMENT,True)
Call TakeScreenSHot()
Call ClickBUtton("Execute   \(Enter\)",True)

CAll VerifyifGuiLabelExists_ByIndex(DT_F02_120_CHECK_TEXT_OF_NO_NAME, 0)
CAll VerifyifGuiLabelExists_ByIndex(DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC2, 0)
CAll VerifyifGuiLabelExists_ByIndex(DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC3, 0)
CAll VerifyifGuiLabelExists_ByIndex(DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC3, 1)

'''''''--------TransactionCode-FBRA----------''''
''
Call SetTcode(DT_EXPECTEDTRANSACTIONCODE_OCC6)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Clearing Document","RF05R-AUGBL","",DT_F02_100_CLEARING_DOCUMENT,FAlse)
Call ClickButton("List Cleared Items   \(Ctrl\+F2\)",False)

Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC7)
Call TakeScreenShot
CAll VerifyifGuiLabelExists_ByIndex(DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC10, 0)
CAll VerifyifGuiLabelExists_ByIndex(DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC11, 0)
CAll VerifyifGuiLabelExists_ByIndex(DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC12, 0)
CAll VerifyifGuiLabelExists_ByIndex(DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC13, 1)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",FAlse)
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC8)
Call ClickButton("Reset cleared items   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call ClickButton("Reset and Reverse",True)
Call SetTextbox("Reversal Reason","RF05R-STGRD","",DT_F02_300_REVERSAL_REASON,True)
Call SetTextbox("Posting Date","RF05R-BUDAT","",ConvertDAte(DT_F02_300_POSTING_DATE),True)
Call SetTextbox("Posting period","RF05R-MONAT","",DT_F02_300_POSTING_PERIOD,True)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot()
Call VerifyTextBoxContent("Information Message","MESSTXT1", 0, Lcase(DT_F02_10_CHECK_TEXT_OF_MESSTXT1_OCC2), True)
Call ClickButton("Continue   \(Enter\)",FAlse)
Call GetTextboxValue("MESSTXT1", 0, "DT_F02_10_CHECK_TEXT_OF_MESSTXT1_OCC3_OUTPUT", True)
Call ClickButtonifexist("Continue   \(Enter\)",True)
Call TakeScreenShot()

'''''''--------TransactionCode-FB03----------''''
'
Call SetTcode(DT_F02_100_OKCD_OCC2)     
Call PressEnter()     
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC9)

Call SetTextbox("Document Number","RF05L-BELNR","",DT_F02_100_DOCUMENT_NUMBER,False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()

Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_F02_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_F02_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

Call VerifyGridCellContent("", 1, "KTONR", 0, DT_F02_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_F02_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)

Call VerifyGridCellContent("", 1, "Amount", 0, DT_F02_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PSWBT)
Call VerifyGridCellContent("", 2, "Amount", 0, DT_F02_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PSWBT)

''''''''--------TransactionCode-FAGLL03----------''''
'
Call SetTcode(DT_F02_750_OKCD)     
Call PressEnter()     
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC10)

Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_F02_1000_GL_ACCOUNT_OCC2,False)
CAll SelectRadioButton("X_AISEL", "All Items", False)
Call SetTextbox("Posting Date","SO_BUDAT-LOW","",ConvertDAte(DT_F02_1000_POSTING_DATE_OCC2),False)
Call SetTextbox("to","SO_BUDAT-HIGH","",ConvertDAte(DT_F02_1000_TO_OCC2),False)
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)

Call ClickLabel("DocumentNo", 0, False)
Call TakeScreenshot()
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call TakeScreenshot()

Call SetTextbox("Document Number","%%DYN001-LOW","",DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC16,True)
Call SetTextbox("to","%%DYN001-HIGH","",DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC17,True)
Call TakeScreenSHot()
Call PressEnter()

CAll VerifyifGuiLabelExists(DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC16)
CAll VerifyifGuiLabelExists(DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC17)

Call ClickLabel("DocumentNo", 0, False)
Call TakeScreenshot()
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call TakeScreenshot()
Call SetTextbox("Document Number","%%DYN001-LOW","","",True)
Call SetTextbox("to","%%DYN001-HIGH","","",True)
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_F02_1105_DOCUMENT_NUMBER,True)
Call TakeScreenSHot()
Call PressEnter()

CAll VerifyifGuiLabelExists_ByIndex(DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC18, 0)
CAll VerifyifGuiLabelExists_ByIndex(DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC19, 1)

Call LogOff()
Call FinalStatus()
