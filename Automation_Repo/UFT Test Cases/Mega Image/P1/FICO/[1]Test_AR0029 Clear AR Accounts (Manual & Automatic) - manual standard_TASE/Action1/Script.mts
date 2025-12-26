'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AR0029 Clear AR Accounts (Manual & Automatic) - manual standard
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


gstrTestCaseName = "Test_AR0029 Clear AR Accounts (Manual & Automatic) - manual standard"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''''----------------------Tcode F-22----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE_F22) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_F22)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F22_100_DOCUMENT_DATE),False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F22_100_COMPANY_CODE,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F22_100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F22_100_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F22_100_DOCHEADER_TEXT,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F22_100_ACCOUNT,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 


Call SetTextbox("Amount","BSEG-WRBTR","",DT_F22_301_AMOUNT,False)
Call SetTextbox("SGL Ind","RF05A-NEWUM","",DT_F22_301_SGL_IND,False)
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_F22_301_ASSIGNMENT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F22_301_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F22_301_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F22_301_ACCOUNT,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 
Call PressEnter() 


Call SetTextbox("Amount","BSEG-WRBTR","",DT_F22_304_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F22_304_TAX_CODE,False)
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_F22_304_ASSIGNMENT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F22_304_TEXT,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot()
Call PressEnter() 

Call TakeScreenShot()
'veryfy sattus bar content
Call GetStatusBar("item1","DT_NUM1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_F22_100_CHECK_TEXT_OF_STATUSBAR)

Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F22_100_DOCUMENT_DATE_OCC2),False)
Call SetTextbox("Type","BKPF-BLART","",DT_F22_100_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F22_100_COMPANY_CODE_OCC2,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F22_100_CURRENCYRATE_OCC2,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F22_100_REFERENCE_OCC2,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F22_100_DOCHEADER_TEXT_OCC2,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F22_100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F22_100_ACCOUNT_OCC2,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F22_301_AMOUNT_OCC2,False)
Call SetTextbox("SGL Ind","RF05A-NEWUM","",DT_F22_301_SGL_IND_OCC2,False)
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_F22_301_ASSIGNMENT_OCC2,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F22_301_TEXT_OCC2,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F22_301_PSTKY_OCC2,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F22_301_ACCOUNT_OCC2,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F22_304_AMOUNT_OCC2,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F22_304_TAX_CODE_OCC2,False)
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_F22_304_ASSIGNMENT_OCC2,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F22_304_TEXT_OCC2,False)
Call SetTextbox("Due On","BSEG-ZFBDT","",ConvertDate(DT_F22_304_DUE_ON),False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 
Call PressEnter() 
Call PressEnter() 


Call GetTextboxValue("BSEG-DMBTR",0,"DT_AMOUNT_IN_LC_OUTPUT",False)

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot()
Call PressEnter() 
Call TakeScreenShot()
''Call Pressenter()
'veryfy sattus bar content
Call GetStatusBar("item1","DT_NUM1_OUTPUT1")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_F22_100_CHECK_TEXT_OF_STATUSBAR_OCC2)

''----------------------Tcode FB03----------------------------

'Enter the Tcode
Call SetTcode(DT_TRANSACTIONCODE_FB03) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_FB03)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Number","RF05L-BELNR","",DT_DOC_NO,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FISCAL_YEAR,False)

Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

Call DoubleClickGuiGridCell("",0, 1, "Posting Key", False)
Call TakeScreenShot
Call GetTextboxValue("BSEG-DMBTR",0,"DT_AMOUNT_IN_LCL_CRRNCY_OUTPUT", False)
Call ClickButton("Back   \(F3\)", False)
'Call GetGridContent("",0,"Amt.in loc.cur.",1,"Posting Key","01","DT_AMOUNT_IN_LCL_CRRNCY_OUTPUT")

Call ClickButton("Back   \(F3\)",False)
wait(2)
Call ClickButton("Back   \(F3\)",False)
wait(2)

'''----------------------Tcode FB03----------------------------
'Enter the Tcode
Call SetTcode(DT_TRANSACTIONCODE_FB03_OCC1) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_FB03_OCC1)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Number","RF05L-BELNR","",DT_F22_100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_COMPANY_CODE_OCC1,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FISCAL_YEAR_OCC1,False)

Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

Call DoubleClickGuiGridCell("",0, 1, "Company code", False)
Call TakeScreenShot
Call GetTextboxValue("BSEG-DMBTR",0,"DT_AMOUNT_IN_LCL_CRRNCY_OUTPUT1", False)
Call ClickButton("Back   \(F3\)", False)
'Call GetGridContent("",0,"Amt.in loc.cur.",1,"Posting Key","01","DT_AMOUNT_IN_LCL_CRRNCY_OUTPUT1")

Call ClickButton("Back   \(F3\)",False)
wait(2)

''''----------------------Tcode FBL5N----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE_FBL5N) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_FBL5N)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_F22_1000_CUSTOMER_ACCOUNT,False)
Call SetTextbox("Company code","DD_BUKRS-LOW","","RO02",False)
Call SelectCheckbox("X_SHBV",0,"ON",False)
Call SelectCheckbox("X_NORM",0,"OFF",False)
Call SelectRadioButton("X_OPSEL","Open items",False)

'Capture the screenshot
Call TakeScreenShot()

'Click on Dynamic Selection
Call ClickButtonIfExist("Dynamic selections   \(Shift\+F4\)",False)

Call ActivateNodeGuiTree("","Documents;Special G/L ind.")

SapGuiSession("transaction:=FBL5N").SapGuiWindow("transaction:=FBL5N").SAPGuiButton("name:=%_%%DYN010_%_APP_%-VALU_PUSH","tooltip:=Multiple selection","index:=10").Highlight
SapGuiSession("transaction:=FBL5N").SapGuiWindow("transaction:=FBL5N").SAPGuiButton("name:=%_%%DYN010_%_APP_%-VALU_PUSH","tooltip:=Multiple selection","index:=10").Click


'set filter criteria
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",1,DT_F22_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",2,DT_F22_3010_TABLECELL_SINGLE_VALUE_1,True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Copy   \(F8\)",True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Wait(2)
Call TakeScreenShot()

Call VerifyifGuiLabelExists(DT_F22_120_CHECK_TEXT_OF_NO_NAME)
Call VerifyifGuiLabelExists(DT_F22_120_CHECK_TEXT_OF_NO_NAME_OCC2)

Call ClickButton("Back   \(F3\)",False)
Wait(2)

Call ClickButton("Back   \(F3\)",False)
Wait(2)

''''----------------------Tcode F-32----------------------------
''Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE_F32) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_F32)
'Capture the screenshot
Call TakeScreenShot()

Call SelectCheckbox("RF05A-XNOPS",0,DT_F22_131_NORMAL_OI,False)
Call SetTextbox("Special G/L Ind","RF05A-AGUMS","",DT_F22_131_SPECIAL_GL_IND,False)
Call SetTextbox("Account","RF05A-AGKON","",DT_F22_131_ACCOUNT,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F22_131_COMPANY_CODE,False)
Call SetTextbox("Currency","BKPF-WAERS","",DT_F22_131_CURRENCY,False)
Call SelectRadioButton("RF05A-XPOS1","Document Number",False)
Call TakeScreenShot()

'Click Process Open Items
Call ClickButton("Process Open Items   \(Shift\+F4\)",False) 

Call SetTextbox("From","RF05A-SEL01","0",DT_DOCUMENT_1,False)
Call SetTextbox("From","RF05A-SEL01","1",DT_DOCUMENT_2,False)

'Click Process Open Items
Call ClickButton("Process Open Items   \(Shift\+F4\)",False) 

Call ClickButton("Post   \(Ctrl\+S\)",False)
'Call PressEnter()

'veryfy sattus bar content
Call GetStatusBar("item1","DT_NUM1_OUTPUT2")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_F22_131_CHECK_TEXT_OF_STATUSBAR)


'''----------------------Tcode FB03----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE_FB03_OCC2) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_FB03_OCC2)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Number","RF05L-BELNR","",DT_F22_100_DOCUMENT_NUMBER_OCC2,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_F22_100_COMPANY_CODE_OCC3,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_F22_100_FISCAL_YEAR,False)

Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Display Document Header   \(F5\)",False)

Call VerifyTextBoxContent("Document type","BKPF-BLART",0,DT_FB03_1710_CHECK_TEXT_OF_DOCUMENT_TYPE,True)

Call ClickButtonIfExist("Continue/Confirm   \(Enter\)",True)


Call VerifyGridCellContent("",1,"Posting Key","",DT_FB03_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",1,"Special G/L ind.","",DT_FB03_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_UMSKZ)
Call VerifyGridCellContent("",2,"Posting Key","",DT_FB03_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OCC2)
Call VerifyGridCellContent("",1,"Amount","",DT_FB03_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("",2,"Amount","",DT_FB03_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PSWBT)
Call DoubleClickGuiGridCell("",0,2, "Posting Key", False)
Call TakeScreenShot
Call GetTextboxValue("BSEG-DMBTR",0,"DT_F22_750_GET_GETCELLVALUE_OF_GRIDCELL_0_DMBTR_Output", False)
'Call VerifyTextBoxContent("Amount in LC","BSEG-DMBTR",0,DT_FB03_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMBTR, False)
Call ClickButton("Back   \(F3\)", False)
'Call GetGridContent("",0,"Amount in LC",1,"Posting Key","01","DT_F22_750_GET_GETCELLVALUE_OF_GRIDCELL_0_DMBTR_Output")
Call ClickButton("Back   \(F3\)",False)
Wait(2)

Call ClickButton("Back   \(F3\)",False)
Wait(2)

'''----------------------Tcode FBL5N----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE_FBL5N_OCC1) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_FBL5N_OCC1)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_F22_1000_CUSTOMER_ACCOUNT_OCC2,False)

Call SelectCheckbox("X_SHBV",0,"ON",False)
Call SelectCheckbox("X_NORM",0,"OFF",False)
Call SelectRadioButton("X_CLSEL","Cleared items",False)

'Click on Dynamic Selection
Call ClickButtonIfExist("Dynamic selections   \(Shift\+F4\)",False)

Call ActivateNodeGuiTree("","Documents;Special G/L ind.")

SapGuiSession("transaction:=FBL5N").SapGuiWindow("transaction:=FBL5N").SAPGuiButton("name:=%_%%DYN010_%_APP_%-VALU_PUSH","tooltip:=Multiple selection","index:=10").Highlight
SapGuiSession("transaction:=FBL5N").SapGuiWindow("transaction:=FBL5N").SAPGuiButton("name:=%_%%DYN010_%_APP_%-VALU_PUSH","tooltip:=Multiple selection","index:=10").Click


'set filter criteria
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",1,DT_F22_3010_TABLECELL_SINGLE_VALUE_0_OCC3,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",2,DT_F22_3010_TABLECELL_SINGLE_VALUE_1_OCC3,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",3,DT_F22_3010_TABLECELL_SINGLE_VALUE_2_OCC2,True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Copy   \(F8\)",True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Wait(2)
Call TakeScreenShot()

Call VerifyStatusBarMessageType("S")

Call VerifyifGuiLabelExists_ByIndex("0,00",1)

Call ClickButton("Back   \(F3\)",False)
wait(2)

Call ClickButton("Back   \(F3\)",False)
wait(2)

'***********************************************************************************************************************************************
'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()
'***********************************************************************************************************************************************

