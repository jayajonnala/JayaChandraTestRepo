
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_GL0003_Post_GL_Account_document_on_both_legders_foreign_currency_TASE
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


gstrTestCaseName = "Test_GL0003_Post_GL_Account_document_on_both_legders_foreign_currency_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

Call Find445CurrentPeriod(DT_TODAY,"DT_PERIOD")


''''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
'''''''''--------TransactionCode-F-02----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(Date),False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F02_100_COMPANY_CODE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_F02_100_TYPE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_100_ACCOUNT,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(Date),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F02_100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F02_100_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F02_100_DOCHEADER_TEXT,False)

Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Text","BSEG-SGTXT","",DT_F02_300_TEXT,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_300_AMOUNT,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_F02_1007_COST_CENTER,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_F02_1007_BUSINESS_AREA,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_300_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_300_ACCOUNT,False)

Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_300_AMOUNT_OCC2,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F02_300_TEXT_OCC2,False)

Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
Call TakeScreenShot

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_DOC_NO_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_NO_OUTPUT&" was posted in company code RO02")

Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_NO_OUTPUT",DT_DOC_NO)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call TakeScreenShot
Call ClickButtonIfExist("Cancel   \(F12\)",True)

Call SelectMenuBar("Document;Display")
Call TakeScreenShot

Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
Call ClickButton("Display Document Header   \(F5\)",False)
Call TakeScreenShot

Call VerifyTextBoxContent("Document type", "BKPF-BLART", "", DT_F02_1710_CHECK_TEXT_OF_DOCUMENT_TYPE, True)
Call GetTextboxValue("BKPF-KURSF", 0, "DT_F02_1710_CHECK_TEXT_OF_EXCHANGE_RATE_OUTPUT", True)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_F02_1710_CHECK_TEXT_OF_EXCHANGE_RATE_OUTPUT",DT_F02_1710_CHECK_TEXT_OF_EXCHANGE_RATE)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call ClickButton("Continue/Confirm   \(Enter\)",True)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_F02_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_F02_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

Call VerifyGridCellContent("", 1, "KTONR", 0, DT_F02_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_F02_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)

Call VerifyGridCellContent("", 1, "Amount", 0, DT_F02_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PSWBT)
Call VerifyGridCellContent("", 2, "Amount", 0, DT_F02_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PSWBT)

Call ClickButton("Document Display: General Ledger View   \(Ctrl\+F9\)",False)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_F02_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OCC2)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_F02_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL_OCC2)

Call VerifyGridCellContent("", 1, "KTONR", 0, DT_F02_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR_OCC2)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_F02_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR_OCC2)

Call VerifyGridCellContent("", 1, "Amount", 0, DT_F02_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PSWBT_OCC2)
Call VerifyGridCellContent("", 2, "Amount", 0, DT_F02_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PSWBT_OCC2)


Call ClickButton("Select Other Ledger   \(Ctrl\+F8\)",False)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_F02_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OCC3)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_F02_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL_OCC3)

Call VerifyGridCellContent("", 1, "KTONR", 0, DT_F02_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR_OCC3)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_F02_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR_OCC3)

Call VerifyGridCellContent("", 1, "Amount", 0, DT_F02_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PSWBT_OCC3)
Call VerifyGridCellContent("", 2, "Amount", 0, DT_F02_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PSWBT_OCC3)


'''''''''--------TransactionCode-ob08----------''''

Call SetTcode(DT_OB08)     
Call PressEnter()     
Call TakeScreenShot

Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call PressEnter()

Call ClickButton("Position\.\.\.",False)
Call TakeScreenShot

Call SetTextbox("Exch\. Rate Type","SVALD-VALUE","",DT_EX_TYPE,True)
Call SetTextbox("From currency","SVALD-VALUE","",DT_FROM,True)
''Call SetTextbox("To-currency","SVALD-VALUE","",DT_FROM,True)
Call SetTextbox("To-currency","SVALD-VALUE","",DT_TO,True)
Call SetTextbox("Valid from","SVALD-VALUE","",ConvertDAte(Date),True)

Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot

'''''''''--------TransactionCode-FAGLL03----------''''

Call SetTcode(DT_F02_100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call ClickButton("Choose Ledger   \(Ctrl\+F4\)",False)
Call TakeScreenShot

Call SetTextbox("Ledger","SVALD-VALUE","",DT_F02_0300_LEDGER,True)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Account Number","RACCT-LOW","",DT_F02_1000_ACCOUNT_NUMBER,False)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
'
Call DoubleClickGuiGridCell("", 0, (Cint(DT_ROW)+1), "Balance", False)

Call TakeScreenShot

Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)

Call ClickButton("Find",True)
Call SetTextbox("Find","GD_SEARCHSTR","","Document Number",True)
Call ClickButton("Continue   \(Enter\)",True)
Call ClickButton("Show sel\. fields \(CTRL\+F3\)",True)
Call ClickButton("Copy   \(Enter\)",True)

Call SetTextbox("Document Number","%%DYN001-LOW","",DT_F02_1105_DOCUMENT_NUMBER,False)
Call ClickButton("Execute   \(Enter\)",True)

Call VerifyifGuiLabelExists(DT_F02_120_CHECK_TEXT_OF_NO_NAME)
Call VerifyifGuiLabelExists(DT_F02_120_CHECK_TEXT_OF_SA)
Call VerifyifGuiLabelExists(DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC2)
Call VerifyifGuiLabelExists(DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC3)
Call VerifyifGuiLabelExists(DT_F02_120_CHECK_TEXT_OF_GL0003BKTXT)

Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot

Call ClickButton("Choose Ledger   \(Ctrl\+F4\)",False)
Call TakeScreenShot

Call SetTextbox("Ledger","SVALD-VALUE",0,DT_F02_0300_LEDGER_OCC2,True)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Account Number","RACCT-LOW","",DT_F02_1000_ACCOUNT_NUMBER,False)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call DoubleClickGuiGridCell("", 0,(DT_ROW+1), "Balance", False)
Call TakeScreenShot

Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)

Call ClickButton("Find",True)
Call SetTextbox("Find","GD_SEARCHSTR","","Document Number",True)
Call ClickButton("Continue   \(Enter\)",True)
Call ClickButton("Show sel\. fields \(CTRL\+F3\)",True)
Call ClickButton("Copy   \(Enter\)",True)

Call SetTextbox("Document Number","%%DYN001-LOW","",DT_F02_1105_DOCUMENT_NUMBER_OCC2,False)
Call ClickButton("Execute   \(Enter\)",True)

Call VerifyifGuiLabelExists(DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC6)
Call VerifyifGuiLabelExists(DT_F02_120_CHECK_TEXT_OF_SA_OCC2)
Call VerifyifGuiLabelExists(DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC7)
Call VerifyifGuiLabelExists(DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC8)
Call VerifyifGuiLabelExists(DT_F02_120_CHECK_TEXT_OF_GL0003BKTXT_OCC2)

Call Logoff'
Call FinalStatus()
