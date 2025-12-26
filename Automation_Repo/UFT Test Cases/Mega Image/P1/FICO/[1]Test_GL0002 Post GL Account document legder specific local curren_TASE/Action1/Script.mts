


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


gstrTestCaseName = "Test_GL0002 Post GL Account document legder specific local curren_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
''

'''----------------------Tcode FAGLB03----------------------------
'
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()   
Call SetTextbox("Account Number","RACCT-LOW","",DT_FAGLB03_1000_ACCOUNT_NUMBER,False)    
Call SetTextbox("Company Code","RBUKRS-LOW","",DT_FAGLB03_1000_COMPANY_CODE,False)    
Call SetTextbox("Fiscal Year","RYEAR","",DT_FAGLB03_1000_FISCAL_YEAR,False)    
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False) 
Call TakeScreenShot()
Call GetGridContent("",0,"Debit",MOnth(Date)+1,"Period",MOnth(Date),"DT_DEBIT_AMNT_OUTPUT")
Call GetGridContent("",0,"Balance",MOnth(Date)+1,"Period",MOnth(Date),"DT_BALANCE_AMNT_OUTPUT")


Call ClickButtonIfExist("Back   \(F3\)",False)
Call ClickButton("Choose Ledger   \(Ctrl\+F4\)",False)
Call TakeScreenShot

Call SetTextbox("Ledger","SVALD-VALUE",0,DT_FAGLB03_0300_LEDGER,True)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False) 
Call TakeScreenShot()

Call GetGridContent("",0,"Debit",MOnth(Date)+1,"Period",MOnth(Date),"DT_DEBIT_AMNT_OUTPUT_1")
Call GetGridContent("",0,"Balance",MOnth(Date)+1,"Period",MOnth(Date),"DT_BALANCE_AMNT_OUTPUT_1")
Call ClickButtonIfExist("Back   \(F3\)",False)
Call SetTextbox("Account Number","RACCT-LOW","",DT_FAGLB03_1000_ACCOUNT_NUMBER_OCC2,False)   
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False) 
Call TakeScreenShot()

Call GetGridContent("",0,"Debit",MOnth(Date)+1,"Period",MOnth(Date),"DT_DEBIT_AMNT_OUTPUT_2")
Call GetGridContent("",0,"Balance",MOnth(Date)+1,"Period",MOnth(Date),"DT_BALANCE_AMNT_OUTPUT_2")
Call ClickButtonIfExist("Back   \(F3\)",False)
Call ClickButton("Choose Ledger   \(Ctrl\+F4\)",False)
Call TakeScreenShot

Call SetTextbox("Ledger","SVALD-VALUE",0,DT_FAGLB03_0300_LEDGER_OCC2,True)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False) 
Call TakeScreenShot()
Call VerifyGridCellContent("",MOnth(Date)+1,"Debit",0,"")

''''------------------------------------------------------------------
Call SetTcode(DT_FAGLB03_100_OKCD) 
Call PressEnter()  

Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_FAGLB03_100_DOCUMENT_DATE),False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FAGLB03_100_COMPANY_CODE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_FAGLB03_100_TYPE,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FAGLB03_100_ACCOUNT,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_FAGLB03_100_DOCUMENT_DATE),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FAGLB03_100_CURRENCYRATE,False)
Call SetTextbox("Ledger Grp","BKPF-LDGRP","",DT_FAGLB03_100_LEDGER_GRP,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_FAGLB03_100_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_FAGLB03_100_DOCHEADER_TEXT,False)
Call TakeScreenShot
Call PressEnter()
Call PressEnter()  
Call TakeScreenShot
Call SetTextbox("Text","BSEG-SGTXT","",DT_FAGLB03_300_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FAGLB03_300_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FAGLB03_300_ACCOUNT,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FAGLB03_300_AMOUNT,False)
Call TakeScreenShot
''Added below two lines to update transaction code
Call ClickButton("COBL_MORE",False) 
Call SetTextbox("Transactn type","COBL-RMVCT","",DT_FAGLB03_300_TRANS_TYPE,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True) 
Call SetTextbox("Text","BSEG-SGTXT","",DT_FAGLB03_300_TEXT_OCC2,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FAGLB03_300_AMOUNT_OCC2,False)
Call TakeScreenShot
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_FAGLB03_1007_COST_CENTER,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_FAGLB03_1007_BUSINESS_AREA,False)
Call TakeScreenShot
'Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_DOC_NO_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_NO_OUTPUT&" was posted in company code RO02")
Call TakeScreenShot

'--------------------------------------------------------------------------------------------------------
Call SetTcode(DT_FAGLB03_100_OKCD_OCC2) 
Call PressEnter() 
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False) 
Call TakeScreenShot
Call VerifyGridCellContent("",MOnth(Date)+1,"Debit",0,"")
Call ClickButtonIfExist("Back   \(F3\)",False)
Call ClickButton("Choose Ledger   \(Ctrl\+F4\)",False)
Call TakeScreenShot

Call SetTextbox("Ledger","SVALD-VALUE",0,DT_FAGLB03_0300_LEDGER_OCC3,True)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False) 
Call TakeScreenShot()
Call GetGridContent("",0,"Debit",MOnth(Date)+1,"Period",MOnth(Date),"DT_DEBIT_AMNT_OUTPUT_4")
Call GetGridContent("",0,"Balance",MOnth(Date)+1,"Period",MOnth(Date),"DT_BALANCE_AMNT_OUTPUT_4")

'---------------------------------------------------------------------------
Call SetTcode(DT_FAGLB03_100_OKCD_OCC3) 
Call PressEnter() 
Call TakeScreenShot
Call PressEnter() 
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_FAGLB03_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_FAGLB03_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
'***********************************************************************************************************************************************
'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()
'***********************************************************************************************************************************************
''''
''''
''''
'''''Enter the Tcode
''''Call SetTcode(DT_SAPTRANSACTIONCODE) 
''''Call PressEnter()     ' 
''''Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
''''
'''''Enter the Details
''''Call SetTextbox("File path name","P_FILE","",DT_ZFIGL_UPLOAD_POST_1000_FILE_PATH_NAME,False)   
''''Call SetTextbox("Session","P_SESS","",DT_ZFIGL_UPLOAD_POST_1000_SESSION,False) 
'''''Capture the screenshot
''''Call TakeScreenShot()
''''
'''''Click execute
''''Call ClickButton("Execute   \(F8\)",False) 
''''
''''
'''''Set ODialog=Dialog("regexpwndtitle:=Microsoft Excel","text:=Microsoft Excel")
'''''ODialog.WinButton("regexpwndtitle:=&Yes","attached text:=There is a large amount of information on the Clipboard.*","text:=&Yes").Highlight
''''''Capture the screenshot
'''''Call TakeScreenShot()
'''''ODialog.WinButton("regexpwndtitle:=&Yes","attached text:=There is a large amount of information on the Clipboard.*","text:=&Yes").Click
'''''Wait(2)
''''
''''
'''''Click execute
''''Call ClickButton("Execute   \(F8\)",False) 
''''Wait(2)
''''
''''
'''''Click execute
''''Call ClickButton("Yes",True) 
''''Wait(2)
''''
'''''Capture the screenshot
''''Call TakeScreenShot()
''''
''''Call ClickButtonIfExist("Back   \(F3\)",False)
''''wait(2)
''''
''''Call ClickButtonIfExist("Back   \(F3\)",False)
''''wait(2)
'''''----------------------Tcode SM35----------------------------
''''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''''
''''Call SetTcode(DT_ZFIGL_UPLOAD_POST_100_OKCD) 
''''Call PressEnter()     ' 
''''Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)
''''Call TakeScreenShot()
''''Wait(2)
''''
''''Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI",1,False)
'''''Click execute
''''Call ClickButton("Process session   \(F8\)",False) 
''''Wait(2)
''''Call SelectRadioButtonByIndexIfPopupExists("D0300-BATCH",0)
''''Call TakeScreenShot()
''''Wait(2)
'''''Click execute
''''Call ClickButton("Process   \(Enter\)",True) 
''''Wait(2)
''''
'''''Verify the status bar message
''''Call VerifyStatusBar(DT_ZFIGL_UPLOAD_POST_1000_CHECK_TEXT_OF_STATUSBAR)
''''
''''
'''''Click on Log
''''Call ClickButton("Log   \(F7\)",False) 
''''Wait(2)
''''
'''''Click on Analyze session and logs
''''Call ClickButton("Analyze session and logs   \(Shift\+F6\)",False) 
''''Wait(2)
''''
'''''Verify the Status
''''Call GetTableCellData("RSBDC_ANALYSETC_TCODES","Status",1,"Trans","FB01","DT_TMP_VAL",False)
''''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''''Call VerifyTableCellContent(1,"Status","RSBDC_ANALYSETC_TCODES",LCase("Processed"))
''''
''''
'''''Navigate to Log Created Tab
''''DT_DATE = Date()
''''
''''Call SelectTab("TAB_DYNPRO"," Log created on "&Replace((DT_DATE),"/","."),False)
''''Wait(1)
''''Call TakeScreenShot()
''''
'''''Verify the message
''''Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",3,"Transaction","FB01","DT_MESSAGE_1",False)
''''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''''Call VerifyTableCellContent(3,"Message","RSBDC_ANALYSETC_PROTOCOL",LCase(DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_8))
''''
''''Call ClickButtonIfExist("Back   \(F3\)",False)
''''wait(2)
''''
''''Call ClickButtonIfExist("Back   \(F3\)",False)
''''wait(2)
''''
''''
'''''----------------------Tcode FB03----------------------------
''''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''''
''''Call SetTcode(DT_ZFIGL_UPLOAD_POST_100_OKCD_OCC2) 
''''Call PressEnter()     ' 
''''Call CheckTCodeScreen(DT_ZFIGL_UPLOAD_POST_100_OKCD_OCC2)
''''Call TakeScreenShot()
''''Wait(2)
''''
'''''Enter the details
''''Call SetTextbox("Document Number","RF05L-BELNR","",DT_ZFIGL_UPLOAD_POST_100_DOCUMENT_NUMBER,False)    
''''Call SetTextbox("Company Code","RF05L-BUKRS","",DT_ZFIGL_UPLOAD_POST_100_COMPANY_CODE,False)    
''''Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_ZFIGL_UPLOAD_POST_100_FISCAL_YEAR,False)    
''''Call TakeScreenShot()
''''Wait(2)
''''Call PressEnter()     ' 
''''
''''
'''''Click Display Document Header   \(F5\)
''''Call ClickButton("Display Document Header   \(F5\)",False) 
''''Wait(2)
''''
'''''Enter the details
''''Call VerifyTextBoxContent("Document type","BKPF-BLART",0,DT_ZFIGL_UPLOAD_POST_1710_CHECK_TEXT_OF_DOCUMENT_TYPE,True)
''''Call VerifyTextBoxContent("Doc\.Header Text","BKPF-BKTXT",0,LCase(DT_ZFIGL_UPLOAD_POST_1710_CHECK_TEXT_OF_DOCHEADER_TEXT),True)
''''Call VerifyTextBoxContent("Reference","BKPF-XBLNR",0,DT_ZFIGL_UPLOAD_POST_1710_CHECK_TEXT_OF_SESSION_NAME,True)
''''Call TakeScreenShot()
''''Wait(2)
''''
''''
'''''Click Continue/Confirm   \(Enter\)
''''Call ClickButton("Continue/Confirm   \(Enter\)",True) 
''''Wait(2)
''''
'''''Select the G/L Account column and press Find
''''Call SelectColumnGuiGrid("",0,"G/L Account",False)
''''Call ClickButtonToolBar("&FIND",0)
''''
''''Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","G/L Account",True) 
''''Call TakeScreenShot()
''''Wait(2)
''''Call ClickButtonIfExist("OK   \(Enter\)",True)
''''Call ClickButtonIfExist("Cancel   \(F12\)",True)
''''
''''
'''''Click Document Display: General Ledger View
''''Call ClickButton("Document Display: General Ledger View   \(Ctrl\+F9\)",False) 
''''Wait(2)
''''Call TakeScreenShot()
''''Wait(2)
''''
''''Call VerifyGridCellContent("",1,"Account",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
''''Call VerifyGridCellContent("",2,"Account",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)
''''Call VerifyGridCellContent("",3,"Account",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_HKONT)
''''Call VerifyGridCellContent("",4,"Account",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_HKONT)
''''Call VerifyGridCellContent("",5,"Account",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_HKONT)
''''Call VerifyGridCellContent("",6,"Account",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_HKONT)
''''Call VerifyGridCellContent("",7,"Account",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_6_HKONT)
''''Call VerifyGridCellContent("",8,"Account",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_7_HKONT)
''''
'''''Verify the Posting Key Values
''''Call VerifyGridCellContent("",1,"Posting Key",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
''''Call VerifyGridCellContent("",2,"Posting Key",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
''''Call VerifyGridCellContent("",3,"Posting Key",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
''''Call VerifyGridCellContent("",4,"Posting Key",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BSCHL)
''''Call VerifyGridCellContent("",5,"Posting Key",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BSCHL)
''''Call VerifyGridCellContent("",6,"Posting Key",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_BSCHL)
''''Call VerifyGridCellContent("",7,"Posting Key",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_6_BSCHL)
''''Call VerifyGridCellContent("",8,"Posting Key",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_7_BSCHL)
''''
'''''Click Select Other Ledger
''''Call ClickButton("Select Other Ledger   \(Ctrl\+F8\)",False) 
''''Wait(2)
''''Call TakeScreenShot()
''''Wait(2)
