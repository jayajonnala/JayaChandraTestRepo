'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AR0030 Clear AR Accounts (Manual & Automatic) - manual residual
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


gstrTestCaseName = "Test_AR0030 Clear AR Accounts (Manual & Automatic) - manual residual"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''----------------------Tcode F-22----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Date","BKPF-BLDAT","",Replace((DT_F22_0100_DOCUMENT_DATE),"/","."),False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F22_0100_COMPANY_CODE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_F22_0100_TYPE,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F22_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F22_0100_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F22_0100_DOCHEADER_TEXT,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F22_0100_ACCOUNT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F22_0100_PSTKY,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 


Call SetTextbox("Amount","BSEG-WRBTR","",DT_F22_0301_AMOUNT,False)
Call SelectCheckbox("BKPF-XMWST",0,"ON",False)
'Call SetTextbox("SGL Ind","RF05A-NEWUM","",DT_F22_301_SGL_IND,False)
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_F22_0301_ASSIGNMENT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F22_0301_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F22_0301_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F22_0301_ACCOUNT,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 
Call PressEnter() 


Call SetTextbox("Amount","BSEG-WRBTR","",DT_F22_0300_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F22_0300_TAX_CODE,False)
Call PressEnter() 
Call SetTextbox("Business Area","COBL-GSBER","",DT_F22_1006_BUSINESS_AREA,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_F22_1006_COST_CENTER,False)
Call SetTextbox("Profit Center","COBL-PRCTR","",DT_F22_1006_PROFIT_CENTER,False)
Call PressEnter() 
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_F22_0300_ASSIGNMENT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F22_0300_TEXT,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()
'veryfy sattus bar content
Call GetStatusBar("item1","DT_NUM1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_F22_0100_CHECK_TEXT_OF_STATUSBAR)

Call SetTextbox("Document Date","BKPF-BLDAT","",Replace((DT_F22_0100_DOCUMENT_DATE_OCC1),"/","."),False)
Call SetTextbox("Type","BKPF-BLART","",DT_F22_0100_TYPE_OCC1,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F22_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F22_0100_CURRENCYRATE_OCC1,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F22_0100_REFERENCE_OCC1,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F22_0100_DOCHEADER_TEXT_OCC1,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F22_0100_PSTKY_OCC1,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F22_0100_ACCOUNT_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F22_0301_AMOUNT_OCC1,False)
'Call SetTextbox("SGL Ind","RF05A-NEWUM","",DT_F22_301_SGL_IND_OCC2,False)
Call SelectCheckbox("BKPF-XMWST",0,"ON",False)
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_F22_0301_ASSIGNMENT_OCC1,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F22_0301_TEXT_OCC1,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F22_0301_PSTKY_OCC1,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F22_0301_ACCOUNT_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 
Call PressEnter()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F22_0300_AMOUNT_OCC1,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F22_0300_TAX_CODE_OCC1,False)
Call PressEnter() 
Call SetTextbox("Business Area","COBL-GSBER","",DT_F22_1006_BUSINESS_AREA_OCC1,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_F22_1006_COST_CENTER_OCC1,False)
Call SetTextbox("Profit Center","COBL-PRCTR","",DT_F22_1006_PROFIT_CENTER_OCC1,False)
Call PressEnter()
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_F22_0300_ASSIGNMENT_OCC1,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F22_0300_TEXT_OCC1,False)
'Call SetTextbox("Due On","BSEG-ZFBDT","",Replace((DT_F22_304_DUE_ON),"/","."),False)

'Capture the screenshot
Call TakeScreenShot()
'Call PressEnter() 
'Call PressEnter() 
'Call PressEnter() 

Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
'Capture the screenshot
Call TakeScreenShot()
'Call GetTextboxValue("BSEG-DMBTR",0,"DT_AMOUNT_IN_LC_OUTPUT",False)

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()
'veryfy sattus bar content
Call GetStatusBar("item1","DT_NUM1_OUTPUT1")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_F22_0100_CHECK_TEXT_OF_STATUSBAR_OCC1)

'''----------------------Tcode FB03----------------------------
'
'Enter the Tcode
Call SetTcode(DT_F22_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Number","RF05L-BELNR","",DT_F22_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_F22_0100_COMPANY_CODE_OCC2,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_F22_0100_FISCAL_YEAR,False)

Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

Call GetTextboxValue("BKPF-XBLNR",0,"DT_F22_0750_CHECK_TEXT_OF_REFERENCE_OUTPUT",False)

'Call GetGridContent("",0,"Amt.in loc.cur.",1,"Posting Key","01","DT_AMOUNT_IN_LCL_CRRNCY_OUTPUT")

Call ClickButton("Back   \(F3\)",False)
wait(2)
Call ClickButton("Back   \(F3\)",False)
wait(2)


'''----------------------Tcode F-32----------------------------
'Enter the Tcode
Call SetTcode(DT_F22_0100_OKCD_OCC1) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)
'Capture the screenshot
Call TakeScreenShot()

Call SelectRadioButton("RF05A-XPOS1","Document Number",False)
'Call SelectCheckbox("RF05A-XNOPS",0,DT_F22_131_NORMAL_OI,False)
'Call SetTextbox("Special G/L Ind","RF05A-AGUMS","",DT_F22_131_SPECIAL_GL_IND,False)
Call SetTextbox("Account","RF05A-AGKON","",DT_F22_0131_ACCOUNT,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F22_0131_COMPANY_CODE,False)
Call SetTextbox("Currency","BKPF-WAERS","",DT_F22_0131_CURRENCY,False)
Call TakeScreenShot()

'Click Process Open Items
Call ClickButton("Process Open Items   \(Shift\+F4\)",False) 

Call SetTextbox("From","RF05A-SEL01","0",DT_F22_0731_FROM,False)
Call SetTextbox("From","RF05A-SEL01","1",DT_F22_0731_FROM_OCC1,False)

'Click Process Open Items
Call ClickButton("Process Open Items   \(Shift\+F4\)",False) 

Call SetTableData("SAPDF05XTC_6102","CashDiscount",1,"","","", False)
Call SetTableData("SAPDF05XTC_6102","CDPer.",1,"","","", False)
Call TakeScreenShot()
Call SelectTab("TS","Res.Items", False)
Call SetTableData("SAPDF05XTC_6106","Residual Items",1,"","",-100, False)
Call TakeScreenShot()
Call PressEnter() 
Call PressEnter()
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call PressEnter()
Call ClickButton("Supplement   \(F5\)",False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F22_0301_TEXT_OCC2,False)
Call PressEnter()
Call ClickButton("Post   \(Ctrl\+S\)",False)
'veryfy sattus bar content
Call GetStatusBar("item1","DT_DOC3_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_F22_0131_CHECK_TEXT_OF_STATUSBAR)

Call ClickButton("Back   \(F3\)",False)
wait(2)
''----------------------Tcode FB03----------------------------

'Enter the Tcode
Call SetTcode(DT_F22_0100_OKCD_OCC2) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC6)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Number","RF05L-BELNR","",DT_F22_0100_DOCUMENT_NUMBER_OCC1,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_F22_0100_COMPANY_CODE_OCC3,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_F22_0100_FISCAL_YEAR_OCC1,False)

Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Display Document Header   \(F5\)",False)

Call VerifyTextBoxContent("Document type","BKPF-BLART",0,DT_F22_1710_CHECK_TEXT_OF_DOCUMENT_TYPE,True)

Call ClickButtonIfExist("Continue/Confirm   \(Enter\)",True)


Call VerifyGridCellContent("",1,"Assignment","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)
Call VerifyGridCellContent("",2,"Assignment","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR)


Call ClickButton("Back   \(F3\)",False)
Wait(2)

Call ClickButton("Back   \(F3\)",False)
Wait(2)

'''----------------------Tcode FBL5N----------------------------
'Enter the Tcode
Call SetTcode(DT_F22_0100_OKCD_OCC3) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC8)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_F22_1000_CUSTOMER_ACCOUNT,False)

'Call SelectCheckbox("X_SHBV",0,"ON",False)
'Call SelectCheckbox("X_NORM",0,"OFF",False)
Call SelectRadioButton("X_AISEL","All items",False)

'Click on Dynamic Selection
Call ClickButtonIfExist("Dynamic selections   \(Shift\+F4\)",False)

Call ActivateNodeGuiTree("","Documents;Special G/L ind.")

SapGuiSession("transaction:=FBL5N").SapGuiWindow("transaction:=FBL5N").SAPGuiButton("name:=%_%%DYN010_%_APP_%-VALU_PUSH","tooltip:=Multiple selection","index:=10").Highlight
SapGuiSession("transaction:=FBL5N").SapGuiWindow("transaction:=FBL5N").SAPGuiButton("name:=%_%%DYN010_%_APP_%-VALU_PUSH","tooltip:=Multiple selection","index:=10").Click


'set filter criteria
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",1,DT_F22_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",2,DT_F22_3010_TABLECELL_SINGLE_VALUE_1,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",3,DT_F22_3010_TABLECELL_SINGLE_VALUE_2,True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Copy   \(F8\)",True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Wait(2)
Call TakeScreenShot()

Call VerifyStatusBarMessageType("S")


Call VerifyifGuiLabelExists_ByIndex("S_LEDR",0)
Call VerifyifGuiLabelExists_ByIndex(DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC5,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC6,0)
Call VerifyifGuiLabelExists_ByIndex("S_LEDG",0)
Call VerifyifGuiLabelExists_ByIndex(DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC9,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F22_0120_CHECK_TEXT_OF_CL,1)
Call VerifyifGuiLabelExists_ByIndex(DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC10,1)
Call VerifyifGuiLabelExists_ByIndex(Replace(DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC7,"/","."),0)
'
Call ClickButton("Change layout...   \(Ctrl\+F8\)", False)
Call SelectRowGuiTableByRow("SAPLSKBHTC_FIELD_LIST",19, True) ' Net due date
Call SelectRowGuiTableByRow("SAPLSKBHTC_FIELD_LIST",72, True) ' Terms of payment
Call ClickButton("Show Sel. Fields \(Ctrl\+F3\)", True)
Call TakeScreenShot
Call ClickButton("Copy   \(Enter\)", True)

Call SetHorizontalScrollBar(100, False)


Call VerifyifGuiLabelExists_ByIndex(DT_F22_0120_CHECK_TEXT_OF_07A1,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F22_0120_CHECK_TEXT_OF_07A1_OCC1,1)
Call VerifyifGuiLabelExists_ByIndex(DT_F22_0120_CHECK_TEXT_OF_07A1_OCC2,2)
Call VerifyifGuiLabelExists_ByIndex(Replace(DT_F22_0120_CHECK_TEXT_OF_NO_NAME,"/","."),0)
Call VerifyifGuiLabelExists_ByIndex(Replace(DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC1,"/","."),0)
Call VerifyifGuiLabelExists_ByIndex(Replace(DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC2,"/","."),0)
Call VerifyifGuiLabelExists_ByIndex(DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC3,0)
'Call ClickButton("Back   \(F3\)",False)
'wait(2)
'
'Call ClickButton("Back   \(F3\)",False)
'wait(2)

'***********************************************************************************************************************************************
'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()
'***********************************************************************************************************************************************

