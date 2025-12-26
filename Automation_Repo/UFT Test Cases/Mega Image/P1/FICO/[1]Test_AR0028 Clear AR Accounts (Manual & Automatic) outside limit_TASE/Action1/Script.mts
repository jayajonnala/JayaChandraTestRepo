
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AR0028 Clear AR Accounts (Manual & Automatic) outside limit
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


gstrTestCaseName = "Test_AR0028"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'----------------------Tcode F-22----------------------------

'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)


Call SetTextbox("Document Date","BKPF-BLDAT","",Replace((DT_F22_0100_DOCUMENT_DATE),"/","."),False)
Call SetTextbox("Type","BKPF-BLART","",DT_F22_0100_TYPE,False)
Call SetTextbox("Period","BKPF-MONAT","",DT_F22_100_PERIOD,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F22_0100_COMPANY_CODE,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F22_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F22_0100_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F22_0100_DOCHEADER_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F22_0100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F22_0100_ACCOUNT,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()


Call SetTextbox("Amount","BSEG-WRBTR","",DT_F22_0301_AMOUNT,False)
Call SetTextbox("Payment Ref\.","BSEG-KIDNO","",DT_F22_0301_PAYMENT_REF,False)
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
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_F22_0301_ASSIGNMENT_OCC1,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F22_0301_TEXT_OCC1,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F22_0301_PSTKY_OCC1,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F22_0301_ACCOUNT_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()
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
Call GetStatusBar("item1","DT_NUM1_OUTPUT1")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_F22_0100_CHECK_TEXT_OF_STATUSBAR_OCC1)

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

Call ClickButton("Yes",True)
Wait(5)

Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)

''----------------------Tcode FB03----------------------------
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

'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()


Call VerifyGridCellContent("",1,"Amount","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("",2,"Amount","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("",1,"Posting Key","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",2,"Posting Key","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call DoubleClickGuiGridCell("",0, 1, "Posting Key", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Acc","BSEG-HKONT",0,DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT, False)
Call ClickButton("Back   \(F3\)", False)
Call DoubleClickGuiGridCell("",0,2, "Posting Key", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Account","BSEG-HKONT",0,DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT, False)
Call ClickButton("Back   \(F3\)", False)
'Call VerifyGridCellContent("",1,"G/L Account","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
'Call VerifyGridCellContent("",2,"G/L Account","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)'

Call ClickButton("Display Another Document   \(Shift\+F5\)",False)
Call TakeScreenShot()
Call SetTextbox("Document Number","RF05L-BELNR","",DT_F22_1110_DOCUMENT_NUMBER,True)
Call ClickButton("Continue/Confirm   \(Enter\)",True)
wait(1)
Call TakeScreenShot()

Call VerifyGridCellContent("",1,"Amount","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET_OCC1)
Call VerifyGridCellContent("",2,"Amount","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET_OCC1)
Call VerifyGridCellContent("",1,"Posting Key","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OCC1)
Call VerifyGridCellContent("",2,"Posting Key","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL_OCC1)
Call DoubleClickGuiGridCell("",0, 1, "Posting Key", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Acc","BSEG-HKONT",0,DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT_OCC1, False)
Call ClickButton("Back   \(F3\)", False)
Call DoubleClickGuiGridCell("",0,2, "Posting Key", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Account","BSEG-HKONT",0,DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT_OCC1, False)
Call ClickButton("Back   \(F3\)", False)
'Call VerifyGridCellContent("",1,"G/L Account","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT_OCC1)
'Call VerifyGridCellContent("",2,"G/L Account","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT_OCC1)'

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)

'''----------------------Tcode F-32----------------------------
'
'Enter the Tcode
Call SetTcode(DT_F22_0100_OKCD_OCC1) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)
'Capture the screenshot
Call TakeScreenShot()


Call SetTextbox("Account","RF05A-AGKON","",DT_F22_0131_ACCOUNT,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F22_0131_COMPANY_CODE,False)
Call SetTextbox("Currency","BKPF-WAERS","",DT_F22_0131_CURRENCY,False)
Call SelectRadioButton("RF05A-XPOS1","Document Number",False)
Call TakeScreenShot()

'Click Process Open Items
Call ClickButton("Process Open Items   \(Shift\+F4\)",False) 

Call SetTextbox("From","RF05A-SEL01","0",DT_F22_0731_FROM,False)
Call SetTextbox("From","RF05A-SEL01","1",DT_F22_0731_FROM_OCC1,False)

'Click Process Open Items
Call ClickButton("Process Open Items   \(Shift\+F4\)",False) 

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call VerifyStatusBar(DT_F22_3100_CHECK_TEXT_OF_STATUSBAR)

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

Call ClickButton("Yes",True)
Wait(2)

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC5)

'***********************************************************************************************************************************************
'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()
'***********************************************************************************************************************************************
