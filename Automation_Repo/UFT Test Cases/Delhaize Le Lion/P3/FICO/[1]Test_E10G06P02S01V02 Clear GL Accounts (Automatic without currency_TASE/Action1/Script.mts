'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_E10G06P02S01V02 Clear GL Accounts (Automatic without currency
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Clear GL Account(AWOC)"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\DLL_P3\Data\TASE_DT_02-04-01-05-03-Create new assortment-BASA.xls"

If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	datatable_row= Parameter("datatable_row")	
End If
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''--------------login----------------'''''

''''Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
'
''--------TransactionCode-F-03----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SelectRadioButton("RF05A-XPOS1","Others", False)
Call TakeScreenShot
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F03_0131_COMPANY_CODE,False)
Call SetTextbox("Account","RF05A-AGKON","",DT_F03_0131_ACCOUNT,False)
Call SetTextbox("Currency","BKPF-WAERS","",DT_F03_0131_CURRENCY,False)
Call SetTextbox("Clearing Date","BKPF-BUDAT","",ConvertDate(DT_F03_0131_CLEARING_DATE),False)
Call SetTextbox("Period","BKPF-MONAT","",DT_F03_0131_PERIOD,False)
Call TakeScreenShot
Call ClickButtonIfExist("Process Open Items   \(Shift\+F4\)",False)
Call TakeScreenShot
'''Call SelectRadioButtonIfPopupExists("RF05A-XPOS1","Collective invoice")
'Call SelectRadioButtonIfExists("RF05A-XPOS1","Collective invoice", True)
'Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call SendKey("{PGDN}")
Call TakeScreenShot
Call SelectRadioButtonIfExists("RF05A-XPOS1","Profit Center",True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SetTextbox("From","RF05A-SEL01","",DT_F03_0731_FROM,False)

Call TakeScreenShot
Call ClickButtonIfExist("Process Open Items   \(Shift\+F4\)",False)
Call TakeScreenShot


Call ClickButton("Select All",False)
Call ClickButton("Deactivate Items",False)
Call TakeScreenShot
Call ClickButton("Field content search",False)
Call TakeScreenShot
Call SelectRadioButton("RF05A-XPOS1","Document Number",True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot
Call SetTextbox("From","RF05A-SEL01","",DT_F03_0731_FROM_OCC1,True)
Call SetTextbox("To","RF05A-SEL02","",DT_F03_0731_TO,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SelectCellGuiTable("SAPDF05XTC_6103","EUR Gross","Document Number",DT_F03_0731_FROM_OCC1,False)
Call SendKey("{F2}")
wait 2
Call SelectCellGuiTable("SAPDF05XTC_6103","EUR Gross","Document Number",DT_F03_0731_TO,False)
Call SendKey("{F2}")
wait 2
Call TakeScreenShot

Call VerifyTextBoxContent("Not assigned","RF05A-DIFFB","",DT_F03_6103_CHECK_TEXT_OF_NOT_ASSIGNED,False)
Call VerifyTableCellContent(1,"Document Number","SAPDF05XTC_6103",DT_F03_6103_CHECK_TEXT_OF_TABLECELL_DOCUMENT_NUMBER_0)
Call VerifyTableCellContent(2,"Document Number","SAPDF05XTC_6103",DT_F03_6103_CHECK_TEXT_OF_TABLECELL_DOCUMENT_NUMBER_1)

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetStatusBar("text", "DT_FAGLB03_0131_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call VerifyStatusBar(DT_F03_0131_CHECK_TEXT_OF_STATUSBAR)

''''''Call PressEnter()     
''''''Call TakeScreenShot()
''''''Call SetTextbox("From","RF05A-SEL01","",DT_FAGLB03_0731_FROM,False)
''''''Call TakeScreenShot
''''''Call PressEnter() 
''''''Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)
''''''Call ClickButtonIfExist("Process Open Items   \(Shift\+F4\)",False)
''''''Call TakeScreenShot
''''''Call VerifyTextBoxContent("Assigned", "RF05A-AKTIV", "", DT_FAGLB03_6103_CHECK_TEXT_OF_NOT_ASSIGNED, False)
''''''Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
''''''Call TakeScreenShot
''''''Call GetStatusBar("item1", "DT_FAGLB03_0131_CHECK_MESSAGEPARAMETER_OF_STATUSBAR")
''''''Call VerifyStatusBar(DT_FAGLB03_0131_CHECK_TEXT_OF_STATUSBAR)


''''''''''--------TransactionCode-FB03----------''''
''
Call SetTcode(DT_EXPECTEDTRANSACTIONCODE_OCC1)     
Call PressEnter()     
Call TakeScreenShot


Call SetTextbox("Document Number","RF05L-BELNR","",DT_F03_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_F03_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_F03_0100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call DoubleClickGuiGridCell("","",1,"Account",False)
Call TakeScreenShot
Call VerifyTextBoxNoLabelContent("SKAT-TXT50","",DT_F03_0300_CHECK_TEXT_OF_GL_ACCOUNT,False)

''''
''''Call SetTextbox("Account Number","RACCT-LOW","",DT_FAGLB03_1000_ACCOUNT_NUMBER,False)
''''Call SetTextbox("Company Code","RBUKRS-LOW","",DT_FAGLB03_1000_COMPANY_CODE,False)
''''Call SetTextbox("Fiscal Year","RYEAR","",DT_FAGLB03_1000_FISCAL_YEAR,False)
''''Call TakeScreenShot
''''Call ClickButton("Execute   \(F8\)",False)
''''Call TakeScreenShot
''''Call DoubleClickGuiGridCell("shell","",6,"Period",False)
''''Call TakeScreenShot
''''Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
''''Call ClickButton("Find",True)
''''Call SetTextbox("Find","GD_SEARCHSTR","","DOCUMENT NUMBER",True)
''''Call ClickButton("Continue   \(Enter\)",True)
''''Call ClickButton("Show sel. fields \(CTRL\+F3\)",True)
''''Call ClickButton("Copy   \(Enter\)",True)
''''Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FAGLB03_1105_DOCUMENT_NUMBER,True)
''''Call TakeScreenShot
''''Call ClickButton("Execute   \(Enter\)",True)
''''Call TakeScreenShot
''''Call VerifyifGuiLabelExists(DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ICO_AUGP)
''''Call TakeScreenShot

'--------TransactionCode-F-03----------''''

Call SetTcode(DT_F03_0300_OKCD)     
Call PressEnter()     
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call SelectRadioButton("RF05A-XPOS1","Others",False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F03_0131_COMPANY_CODE_OCC1,False)
Call SetTextbox("Account","RF05A-AGKON","",DT_F03_0131_ACCOUNT_OCC1,False)
Call SetTextbox("Currency","BKPF-WAERS","",DT_F03_0131_CURRENCY_OCC1,False)
Call SetTextbox("Period","BKPF-MONAT","",DT_F03_0131_PERIOD_OCC1,False)
Call SetTextbox("Clearing Date","BKPF-BUDAT","",ConvertDate(DT_F03_0131_CLEARING_DATE_OCC1),False)
Call TakeScreenShot
Call ClickButtonIfExist("Process Open Items   \(Shift\+F4\)",False)
Call TakeScreenShot
Call SendKey("{PGDN}")
Call TakeScreenShot
Call SelectRadioButton("RF05A-XPOS1","Profit Center",True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SetTextbox("From","RF05A-SEL01","",DT_F03_0731_FROM_OCC2,False)
Call TakeScreenShot
Call ClickButtonIfExist("Process Open Items   \(Shift\+F4\)",False)
Call TakeScreenShot

Call ClickButton("Select All",False)
Call ClickButton("Deactivate Items",False)
Call ClickButton("Field content search",False)
Call SelectRadioButton("RF05A-XPOS1","Document Number",True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SetTextbox("From","RF05A-SEL01","",DT_F03_0731_FROM_OCC3,True)
Call SetTextbox("To","RF05A-SEL02","",DT_F03_0731_TO_OCC1,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SelectCellGuiTable("SAPDF05XTC_6103","EUR Gross","Document Number",DT_F03_0731_FROM_OCC3,False)
Call SendKey("{F2}")
wait 2
Call TakeScreenShot
Call SelectCellGuiTable("SAPDF05XTC_6103","EUR Gross","Document Number",DT_F03_0731_TO_OCC1,False)
Call SendKey("{F2}")
wait 2
Call TakeScreenShot

Call VerifyTextBoxContent("Not assigned","RF05A-DIFFB","",DT_F03_6103_CHECK_TEXT_OF_NOT_ASSIGNED_OCC1,False)
Call VerifyTableCellContent(1,"Document Number","SAPDF05XTC_6103",DT_F03_6103_CHECK_TEXT_OF_TABLECELL_DOCUMENT_NUMBER_0_OCC1)
Call VerifyTableCellContent(2,"Document Number","SAPDF05XTC_6103",DT_F03_6103_CHECK_TEXT_OF_TABLECELL_DOCUMENT_NUMBER_1_OCC1)
Call TakeScreenShot

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetStatusBar("text", "DT_F03_0131_CHECK_TEXT_OF_STATUSBAR_OCC1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call VerifyStatusBar(DT_F03_0131_CHECK_TEXT_OF_STATUSBAR_OCC1)

'''Call SelectRadioButton("RF05A-XPOS1","Document Number", False)
'''Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FAGLB03_0131_COMPANY_CODE,False)
'''Call SetTextbox("Account","RF05A-AGKON","",DT_FAGLB03_0131_ACCOUNT,False)
'''Call SetTextbox("Currency","BKPF-WAERS","",DT_FAGLB03_0131_CURRENCY,False)
'''Call TakeScreenShot
'''Call PressEnter()     
'''Call TakeScreenShot()
'''Call SetTextbox("From","RF05A-SEL01","",DT_FAGLB03_0731_FROM,False)
'''Call TakeScreenShot
'''Call PressEnter() 
'''Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)
'''Call ClickButtonIfExist("Process Open Items   \(Shift\+F4\)",False)
'''Call TakeScreenShot
'''Call VerifyTextBoxContent("Assigned", "RF05A-AKTIV", "", DT_FAGLB03_6103_CHECK_TEXT_OF_NOT_ASSIGNED, False)
'''Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
'''Call TakeScreenShot
'''Call GetStatusBar("item1", "DT_FAGLB03_0131_CHECK_MESSAGEPARAMETER_OF_STATUSBAR")
'''Call VerifyStatusBar(DT_FAGLB03_0131_CHECK_TEXT_OF_STATUSBAR)

''''''--------TransactionCode-FB03---------
'
Call SetTcode(DT_F03_0131_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Document Number","RF05L-BELNR","",DT_F03_0100_DOCUMENT_NUMBER_OCC1,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_F03_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_F03_0100_FISCAL_YEAR_OCC1,False)
Call TakeScreenShot
Call PressEnter()

Call DoubleClickGuiGridCell("","",1,"Account",False)
Call TakeScreenShot

Call VerifyTextBoxNoLabelContent("SKAT-TXT50","",DT_F03_0300_CHECK_TEXT_OF_GL_ACCOUNT_OCC1,False)

''''Call SetTextbox("Account Number","RACCT-LOW","",DT_FAGLB03_1000_ACCOUNT_NUMBER_OCC1,False)
''''Call SetTextbox("Company Code","RBUKRS-LOW","",DT_FAGLB03_1000_COMPANY_CODE_OCC1,False)
''''Call SetTextbox("Fiscal Year","RYEAR","",DT_FAGLB03_1000_FISCAL_YEAR_OCC1,False)
''''Call TakeScreenShot
''''Call ClickButton("Execute   \(F8\)",False)
''''Call TakeScreenShot
''''Call DoubleClickGuiGridCell("shell","",6,"Period",False)
''''Call TakeScreenShot
''''Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
''''Call ClickButton("Find",True)
''''Call SetTextbox("Find","GD_SEARCHSTR","","DOCUMENT NUMBER",True)
''''Call ClickButton("Continue   \(Enter\)",True)
''''Call ClickButton("Show sel. fields \(CTRL\+F3\)",True)
''''Call ClickButton("Copy   \(Enter\)",True)
'''''Call SetTextbox("Document Number","%%DYN001-LOW","",DT_F13_1105_DOCUMENT_NUMBER,True)
'''''Call TakeScreenShot
'''''Call ClickButton("Execute   \(Enter\)",True)
'''''Call TakeScreenShot
'''''Call VerifyifGuiLabelExists(DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ICO_AUGP_OCC1)
'''''Call VerifyifGuiLabelExists(DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
'''''Call VerifyifGuiLabelExists(DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
'''''Call VerifyifGuiLabelExists(DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB)
'''''Call VerifyifGuiLabelExists(DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AUGBL) 
'''''Call TakeScreenShot

''''----------------------FAGLB03-----------------------------
Call SetTcode(DT_F03_0300_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)

Call SetTextbox("Account Number","RACCT-LOW","",DT_F03_1000_ACCOUNT_NUMBER,False)
Call SetTextbox("Company Code","RBUKRS-LOW","",DT_F03_1000_COMPANY_CODE,False)
Call Settextbox("Fiscal Year","RYEAR","",DT_F03_1000_FISCAL_YEAR,False)
Call TakeScreenShot

Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call DoubleClickGuiGridCell("","",5,"Balance",False)'DT_F03_0030_GRIDCELL_13_BALANCE
Call TakeScreenShot

Call SelectColumnGuiGrid("", 0, "Document Number", False)
Call TakeScreenShot
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call TakeScreenShot
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_F03_1105_DOCUMENT_NUMBER,True)
Call TakeScreenShot
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot


''''Call ClickButton("Find",True)
''''Call SetTextbox("Find","GD_SEARCHSTR","",DT_F03_0850_FIND,True)
''''Call ClickButton("Continue   \(Enter\)",True)
''''Call ClickButton("Show sel\. fields \(CTRL\+F3\)",False)
''''Call ClickButton("Copy   \(Enter\)",False)
''''Call SetTextbox("Document Number","%%DYN001-LOW","",DT_F03_1105_DOCUMENT_NUMBER,True)
''''Call ClickButton("Execute   \(Enter\)",True)

'Call Verifygridcellcontent("",1,"Cleared/open items symbol",0,"@5C\QOpen@")
'''Call VerifyifGuiLabelExists_ByIndex(DT_F03_0120_CHECK_TOOLTIP_OF_NO_NAME,0)


''---------------------------FB08--------------------------

Call SetTcode(DT_F03_0120_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC5)

Call SetTextbox("Posting period","BSIS-MONAT","",DT_F03_0105_POSTING_PERIOD,False)
Call SetTextbox("Posting Date","BSIS-BUDAT","",DT_F03_0105_POSTING_DATE,False)
Call SetTextbox("Reversal Reason","UF05A-STGRD","",DT_F03_0105_REVERSAL_REASON,False)
Call SetTextbox("Fiscal Year","RF05A-GJAHS","",DT_F03_0105_FISCAL_YEAR,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F03_0105_COMPANY_CODE,False)
Call SetTextbox("Document Number","RF05A-BELNS","",DT_F03_0105_DOCUMENT_NUMBER,False)

Call ClickButton("Display document before reversal   \(F5\)",False)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC6)
Call VerifyTextBoxContent("Reference","BKPF-XBLNR","",DT_F03_0750_CHECK_TEXT_OF_REFERENCE,False)
Call ClickButton("Back   \(F3\)",False)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC5)
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("text","DT_F03_0105_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)


Call SetTextbox("Posting period","BSIS-MONAT","",DT_F03_0105_POSTING_PERIOD,False)
Call SetTextbox("Fiscal Year","RF05A-GJAHS","",DT_F03_0105_FISCAL_YEAR,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F03_0105_COMPANY_CODE,False)
Call SetTextbox("Document Number","RF05A-BELNS","",DT_F03_0105_DOCUMENT_NUMBER,False)
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call VerifyStatusBar(DT_F03_0105_CHECK_TEXT_OF_STATUSBAR_OCC1)
'
''''----------------------------FAGLB03-------------------------
Call SetTcode(DT_F03_0105_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC8)

Call SetTextbox("Account Number","RACCT-LOW","",DT_F03_1000_ACCOUNT_NUMBER_OCC1,False)
Call SetTextbox("Company Code","RBUKRS-LOW","",DT_F03_1000_COMPANY_CODE_OCC1,False)
Call Settextbox("Fiscal Year","RYEAR","",DT_F03_1000_FISCAL_YEAR_OCC1,False)

Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call DoubleClickGuiGridCell("","",5,"Balance",False)'DT_F03_0030_GRIDCELL_13_BALANCE
Call TakeScreenShot

Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call TakeScreenShot
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot

Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_F03_0850_FIND,True)
Call TakeScreenShot
Call ClickButton("OK   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Cancel   \(F12\)",True)
Call TakeScreenShot
Call ClickButton("Add Filter Criterion \(F7\)",True)
Call TakeScreenShot
Call ClickButton("Define Filter Values",True)
Call TakeScreenShot
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_F03_1105_DOCUMENT_NUMBER_OCC1,True)
Call SetTextbox("to","%%DYN001-HIGH","",DT_F03_1105_TO,True)
Call TakeScreenShot
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot



'
'
''
''Call ClickButton("Transfer   \(Enter\)",True)
''Call TakeScreenShot
''Call ClickButton("Show sel\. fields \(CTRL\+F3\)",False)
''Call TakeScreenShot
''Call ClickButton("Copy   \(Enter\)",False)
''Call TakeScreenShot
''Call ClickButton("Multiple selection",True)
''Call TakeScreenShot
''Call SelectTab("TAB_STRIP","Select Ranges",True)
''Call TakeScreenShot
''Call SetTableData("SAPLALDBINTERVAL","Lower limit",1,"","",DT_F03_1105_DOCUMENT_NUMBER_OCC1,True)
''Call SetTableData("SAPLALDBINTERVAL","Upper limit",1,"","",DT_F03_1105_TO,True)
''Call TakeScreenShot
''Call ClickButton("Copy   \(F8\)",True)
''Call TakeScreenShot
''Call ClickButton("Execute   \(Enter\)",True)
''Call TakeScreenShot
'
Call Verifygridcellcontent("",1,"Cleared/open items symbol",0,"@5B\QCleared@")
Call Verifygridcellcontent("",2,"Cleared/open items symbol",0,"@5B\QCleared@")
Call Verifygridcellcontent("",1,"Document Number",0,DT_F03_0120_CHECK_TEXT_OF_NO_NAME)
Call Verifygridcellcontent("",2,"Document Number",0,DT_F03_0120_CHECK_TEXT_OF_NO_NAME_OCC1)

'
'Call VerifyifGuiLabelExists_ByIndex(DT_F03_0120_CHECK_TOOLTIP_OF_NO_NAME_OCC1,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_F03_0120_CHECK_TOOLTIP_OF_NO_NAME_OCC2,1)
'Call VerifyifGuiLabelExists_ByIndex(DT_F03_0120_CHECK_TEXT_OF_NO_NAME,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_F03_0120_CHECK_TEXT_OF_NO_NAME_OCC1,0)
'''
'''''''-------------------------------------------------------------------
''''Call SetFocusGuiLabel(DT_F03_0120_CHECK_TEXT_OF_NO_NAME,46,136,False)
''''Call VerifyifGuiLabelExists(DT_F03_0120_CHECK_TOOLTIP_OF_NO_NAME_OCC1)
''''Call SetFocusGuiLabel(DT_F03_0120_CHECK_TEXT_OF_NO_NAME,46,152,False)
''''Call VerifyifGuiLabelExists(DT_F03_0120_CHECK_TOOLTIP_OF_NO_NAME_OCC2)
''''Call SetFocusGuiLabel(DT_F03_0120_CHECK_TEXT_OF_NO_NAME,200,136,False)
''''Call VerifyifGuiLabelExists(DT_F03_0120_CHECK_TEXT_OF_NO_NAME)
''''Call SetFocusGuiLabel(DT_F03_0120_CHECK_TEXT_OF_NO_NAME_OCC1,200,152,False)
''''Call VerifyifGuiLabelExists(DT_F03_0120_CHECK_TEXT_OF_NO_NAME_OCC1)
''''
'''''''''-------------------------------------------------------------------
'''''Call CheckifGuiLabelExists(DT_F03_0120_CHECK_TOOLTIP_OF_NO_NAME_OCC1)
'''''Call CheckifGuiLabelExists(DT_F03_0120_CHECK_TOOLTIP_OF_NO_NAME_OCC2)
'''''Call CheckifGuiLabelExists(DT_F03_0120_CHECK_TEXT_OF_NO_NAME)
'''''Call CheckifGuiLabelExists(DT_F03_0120_CHECK_TEXT_OF_NO_NAME_OCC1)
'''
''''''------------------------------------------------------------------------------------------------------------
''''Call VerifyifGuiLabelExistsByRelativeid(DT_F03_0120_CHECK_TOOLTIP_OF_NO_NAME_OCC1,"wnd\[0\]/usr/lbl\[6,8\]")
''''Call VerifyifGuiLabelExistsByRelativeid(DT_F03_0120_CHECK_TOOLTIP_OF_NO_NAME_OCC2,"wnd\[0\]/usr/lbl\[6,9\]")
''''Call VerifyifGuiLabelExistsByRelativeid(DT_F03_0120_CHECK_TEXT_OF_NO_NAME,"wwnd\[0\]/usr/lbl\[28,8\]")
''''Call VerifyifGuiLabelExistsByRelativeid(DT_F03_0120_CHECK_TEXT_OF_NO_NAME_OCC1,"wwnd\[0\]/usr/lbl\[28,9\]")
'
'''------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

''*********************************************End Of Script*********************************************************************
