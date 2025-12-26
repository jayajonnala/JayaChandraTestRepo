
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_GL0016 Upload GL Document with vendor account on both ledger
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


gstrTestCaseName = "Test_GL0016 Upload GL Document with vendor account on both ledger"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call Find445CurrentPeriod(DT_TODAY,"DT_PERIOD")


'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''''----------------------Tcode FAGLB03----------------------------
''
''Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
Call TakeScreenShot()
'
'Verify Ledger
Call  VerifyTextBoxContent("Ledger","RLDNR",0,DT_FAGLB03_1000_CHECK_TEXT_OF_LEDGER,False)
'
'Enter the details
Call SetTextbox("Account Number","RACCT-LOW","",DT_FAGLB03_1000_ACCOUNT_NUMBER,False)    
Call SetTextbox("Company Code","RBUKRS-LOW","",DT_FAGLB03_1000_COMPANY_CODE,False)    
Call SetTextbox("Fiscal Year","RYEAR","",DT_FAGLB03_1000_FISCAL_YEAR,False)    
Call TakeScreenShot()

'Click execute
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()

''Call GetGridContentByTitle("", 0, "Debit", (Month(Date)+1), "DT_DEBIT_AMNT")
''Call GetGridContentByTitle("",0,"Balance",(MOnth(Date)+1),"DT_BALANCE_AMNT")
Call GetGridContentByTitle("", 0, "Debit", (Cint(DT_ROW)+1), "DT_DEBIT_AMNT")
Call GetGridContentByTitle("",0,"Balance",(Cint(DT_ROW)+1),"DT_BALANCE_AMNT")

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)
Call TakeScreenShot()

'Click Choose Ledger
Call ClickButton("Choose Ledger   \(Ctrl\+F4\)",False) 
Call TakeScreenShot()
'
'Select Ledger
Call SetTextbox("Ledger","SVALD-VALUE","",DT_FAGLB03_0300_LEDGER,True)    
Call TakeScreenShot()
Call ClickButton("Cont\.   \(Enter\)",True)
Call TakeScreenShot()
Call PressEnter() 
Call TakeScreenShot()

'Click execute
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()

''
''Call GetGridContentByTitle("",0,"Debit",(MOnth(Date)+1),"DT_DEBIT_AMNT_1")
''Call GetGridContentByTitle("",0,"Balance",(MOnth(Date)+1),"DT_BALANCE_AMNT_1")
Call GetGridContentByTitle("",0,"Debit",(Cint(DT_ROW)+1),"DT_DEBIT_AMNT_1")
Call GetGridContentByTitle("",0,"Balance",(Cint(DT_ROW)+1),"DT_BALANCE_AMNT_1")

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(1)
Call TakeScreenShot()
Call ClickButtonIfExist("Back   \(F3\)",False)
wait(1)
Call TakeScreenShot()

Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)
''
'
'----------------------Tcode ZFIGL_UPLOAD_POST----------------------------
'
'''''Enter the Tcode
Call SetTcode(DT_FAGLB03_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
Call TakeScreenShot()

'''Enter the Details
Call SetTextbox("File path name","P_FILE","",DT_FAGLB03_1000_FILE_PATH_NAME,False)   
Call SetTextbox("Session","P_SESS","",DT_FAGLB03_1005_SESS,False) 
Call TakeScreenShot()

''''Click execute
Call ClickButton("Execute   \(F8\)",False) 
Call TakeScreenShot()

Set ODialog=Dialog("regexpwndtitle:=Microsoft Excel","text:=Microsoft Excel")
If ODialog.WinButton("regexpwndtitle:=&Yes","attached text:=There is a large amount of information on the Clipboard.*","text:=&Yes").Exist(5) Then
	ODialog.WinButton("regexpwndtitle:=&Yes","attached text:=There is a large amount of information on the Clipboard.*","text:=&Yes").Highlight
'''	Capture the screenshot
	Call TakeScreenShot()
	ODialog.WinButton("regexpwndtitle:=&Yes","attached text:=There is a large amount of information on the Clipboard.*","text:=&Yes").Click
	Wait(2)
	
End If

''''''Verify the Grid Content
Call VerifyGridCellContent("",1,"Document Header Text",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BKTXT)
Call VerifyGridCellContent("",1,"Posting Key",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_NEWBS)
Call VerifyGridCellContent("",2,"Posting Key",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NEWBS)
Call VerifyGridCellContent("",3,"Posting Key",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_NEWBS)
Call VerifyGridCellContent("",4,"Posting Key",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_NEWBS)
Call VerifyGridCellContent("",1,"Amount",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WRBTR)
Call VerifyGridCellContent("",2,"Amount",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_WRBTR)
Call VerifyGridCellContent("",3,"Amount",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_WRBTR)
Call VerifyGridCellContent("",4,"Amount",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_WRBTR)
Call VerifyGridCellContent("",1,"Account",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_NEWKO)
Call VerifyGridCellContent("",2,"Account",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NEWKO)
Call VerifyGridCellContent("",3,"Account",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_NEWKO)
Call VerifyGridCellContent("",4,"Account",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_NEWKO)

Call ClickButton("Execute   \(F8\)",False) 
Wait(1)
Call TakeScreenShot()

''''''Click execute
Call ClickButton("Yes",True) 
Wait(1)
'''''Capture the screenshot
Call TakeScreenShot()

Call SelectMenuBar("System;Services;Batch Input;Sessions")
Wait 5 
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)
Call TakeScreenShot()
Wait(2)

Call SetTextbox("Sess\.:","D0100-MAPN","",DT_FAGLB03_1005_SESS,False)    
Call SetTextbox("From:","D0100-VON","",ConvertDate(DT_FAGLB03_1005_FROM),False)    
Call SetTextbox("Created by:","D0100-CREATOR","",DT_FAGLB03_1005_CREATED_BY,False)    
Call TakeScreenShot()
Wait(1)
Call PressEnter()
Wait(3)
Call TakeScreenShot()

Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI",1,False)
Call TakeScreenShot()

Call ClickButton("Process session   \(F8\)",False) 
Wait(2)
Call TakeScreenShot()
Call SelectRadioButtonByIndexIfPopupExists("D0300-BATCH",0)
Call TakeScreenShot()
Wait(2)

Call ClickButton("Process   \(Enter\)",True) 
Wait(5)
Call TakeScreenShot()

'Verify the status bar message
Call VerifyStatusBar(DT_FAGLB03_1000_CHECK_TEXT_OF_STATUSBAR)


Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI",1,False)
Call ClickButton("Analyze session   \(F2\)",False)
Call TakeScreenShot()

'Call SendKey("{F2}")
'wait 5
'Set objWsh = CreateObject("WScript.Shell") 
'objWsh.SendKeys "{TAB}"
'objWsh.SendKeys "{F2}"
'Set objWsh=nothing
''Call SendKey("{F2}")

Call VerifyTableCellContent(1,"Status","RSBDC_ANALYSETC_TCODES",LCase("Processed"))

'Navigate to Log Created Tab
Call SelectTab("TAB_DYNPRO"," Log created on "&ConvertDate(Date),False)
Wait(1)
Call TakeScreenShot()

'Verify the message
''Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",5,"","","DT_MESSAGE_1",False)
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",4,"","","DT_MESSAGE_1",False)

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''Call VerifyTableCellContent(5,"Message","RSBDC_ANALYSETC_PROTOCOL",LCase(DT_MESSAGE_1))
Call VerifyTableCellContent(4,"Message","RSBDC_ANALYSETC_PROTOCOL",LCase(DT_MESSAGE_1))
Call ClickButtonIfExist("Back   \(F3\)",False)
wait(1)
Call TakeScreenShot()
Call ClickButtonIfExist("Back   \(F3\)",False)
wait(1)
Call TakeScreenShot()

Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC8)

'''----------------------Tcode FAGLB03----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'Enter the Tcode
Call SetTcode(DT_FAGLB03_0100_OKCD_OCC3) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC9)
Call TakeScreenShot()


'Enter the details
Call SetTextbox("Account Number","RACCT-LOW","",DT_FAGLB03_1000_ACCOUNT_NUMBER_OCC1,False)    
Call SetTextbox("Company Code","RBUKRS-LOW","",DT_FAGLB03_1000_COMPANY_CODE_OCC1,False)    
Call SetTextbox("Fiscal Year","RYEAR","",DT_FAGLB03_1000_FISCAL_YEAR_OCC1,False)    
Call TakeScreenShot()
Wait(2)

'Click execute
Call ClickButton("Execute   \(F8\)",False) 
Call TakeScreenShot()

''Call GetGridContentByTitle("",0,"Debit",(MOnth(Date)+1),"DT_DEBIT_AMNT_2")
''Call GetGridContentByTitle("",0,"Balance",(MOnth(Date)+1),"DT_BALANCE_AMNT_2")
Call GetGridContentByTitle("",0,"Debit",(Cint(DT_ROW)+1),"DT_DEBIT_AMNT_2")
Call GetGridContentByTitle("",0,"Balance",(Cint(DT_ROW)+1),"DT_BALANCE_AMNT_2")

''Call DoubleClickGuiGridCell("", 0, (Month(Date)+1), "Debit", False)
Call DoubleClickGuiGridCell("", 0, (Cint(DT_ROW)+1), "Debit", False)
Call TakeScreenShot()

''Set Filter
'Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
'Call TakeScreenShot()
'
'Call SelectRowGuiTable("SAPLSKBHTC_FIELD_LIST_820","Content","Document Number",True)
'Call TakeScreenShot()
'Call ClickButton("Show sel\. fields \(CTRL\+F3\)",True)
'Call TakeScreenShot()
'Call ClickButton("Copy   \(Enter\)",True)
'Call TakeScreenShot()
'
'Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FAGLB03_1105_DOCUMENT_NUMBER,True) 
'Call TakeScreenShot()
'Call ClickButton("Execute   \(Enter\)",True)
'Wait(1)
'Call TakeScreenShot()
'
Call VerifyifGuiLabelExistsByRelativeid(DT_FAGLB03_0120_CHECK_TEXT_OF_NO_NAME, "wnd\[0\]/usr/lbl\[43,8\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_FAGLB03_0120_CHECK_TEXT_OF_NO_NAME_OCC1, "wnd\[0\]/usr/lbl\[43,9\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_FAGLB03_0120_CHECK_TEXT_OF_NO_NAME_OCC2, "wnd\[0\]/usr/lbl\[43,10\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_FAGLB03_0120_CHECK_TEXT_OF_NO_NAME_OCC3, "wnd\[0\]/usr/lbl\[43,11\]")

''Call VerifyifGuiLabelExistsByRelativeid(DT_FAGLB03_0120_CHECK_TEXT_OF_OTH_EXP_NONDED, "wnd\[0\]/usr/lbl\[9,8\]")

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(10)
Call TakeScreenShot()
''Call ClickButtonIfExist("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)
wait(10)
Call TakeScreenShot()

''Click Choose Ledger
Call ClickButton("Choose Ledger   \(Ctrl\+F4\)",False) 
Call TakeScreenShot()

'Select Ledger
Call SetTextbox("Ledger","SVALD-VALUE","",DT_FAGLB03_0300_LEDGER_OCC1,True)    
Call TakeScreenShot()
Call ClickButton("Cont\.   \(Enter\)",True)
Call PressEnter() 
Call TakeScreenShot()

'Click execute
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()


''Call GetGridContentByTitle("",0,"Debit",(MOnth(Date)+1),"DT_DEBIT_AMNT_3")
''Call GetGridContentByTitle("",0,"Balance",(MOnth(Date)+1),"DT_BALANCE_AMNT_3")
Call GetGridContentByTitle("",0,"Debit",(Cint(DT_ROW)+1),"DT_DEBIT_AMNT_3")
Call GetGridContentByTitle("",0,"Balance",(Cint(DT_ROW)+1),"DT_BALANCE_AMNT_3")

Wait 5

''Call DoubleClickGuiGridCell("", 0, (Month(Date)+1), "Debit", False)
Call DoubleClickGuiGridCell("", 0, (Cint(DT_ROW)+1), "Debit", False)

Call TakeScreenShot()

'Set Filter
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call TakeScreenShot()

Call SelectRowGuiTable("SAPLSKBHTC_FIELD_LIST_820","Content","Document Number",True)
Call TakeScreenShot()
Call ClickButton("Show sel\. fields \(CTRL\+F3\)",True)
Call TakeScreenShot()
Call ClickButton("Copy   \(Enter\)",True)
Call TakeScreenShot()

Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FAGLB03_1105_DOCUMENT_NUMBER_OCC1,True) 
Call TakeScreenShot()
Call ClickButton("Execute   \(Enter\)",True)
Wait(1)
Call TakeScreenShot()
'script change needed here
Call VerifyifGuiLabelExistsByRelativeid(DT_FAGLB03_0120_CHECK_TEXT_OF_NO_NAME_OCC4, "wnd\[0\]/usr/lbl\[43,8\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_FAGLB03_0120_CHECK_TEXT_OF_NO_NAME_OCC5, "wnd\[0\]/usr/lbl\[43,9\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_FAGLB03_0120_CHECK_TEXT_OF_NO_NAME_OCC6, "wnd\[0\]/usr/lbl\[43,10\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_FAGLB03_0120_CHECK_TEXT_OF_NO_NAME_OCC7, "wnd\[0\]/usr/lbl\[43,11\]")

'Call VerifyifGuiLabelExistsByRelativeid(DT_FAGLB03_0120_CHECK_TEXT_OF_OTH_EXP_NONDED_OCC1, "wnd\[0\]/usr/lbl\[9,8\]")


Call ClickButtonIfExist("Back   \(F3\)",False)
wait(10)
Call TakeScreenShot()
Call ClickButtonIfExist("Back   \(F3\)",False)
wait(10)
Call TakeScreenShot()
Call ClickButtonIfExist("Back   \(F3\)",False)
wait(10)
Call TakeScreenShot()
'

''''''----------------------Tcode FBL3N----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

''''Enter the Tcode
Call SetTcode(DT_FAGLB03_0100_OKCD_OCC4) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC11)
Call TakeScreenShot()

Call SelectRadioButton("X_AISEL", "All items", False)
Call TakeScreenShot()

Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_FAGLB03_1000_GL_ACCOUNT,False)  
Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_FAGLB03_1000_COMPANY_CODE_OCC2,False)  
Call SetTextbox("Posting date","SO_BUDAT-LOW","",ConvertDAte(Date),False)  
Call SetTextbox("to","SO_BUDAT-HIGH","",ConvertDAte(Date),False)  
Call TakeScreenShot()

Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call TakeScreenShot()

Call SelectNodeGuiTree(0, "Document;Document Number")
Wait 2
Call TakeScreenShot()
CAll ActivateNodeGuiTree(0, "Document;Document Number")
Wait 2
Call TakeScreenShot()

Call SetTextbox("Document Number","%%DYN007-LOW","",DT_FAGLB03_1106_DOCUMENT_NUMBER,False)  

Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenSHot()

Call VerifyifGuiLabelExistsByRelativeid(DT_FAGLB03_0120_CHECK_TEXT_OF_NO_NAME_OCC9, "wnd\[0\]/usr/lbl\[67,7\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_FAGLB03_0120_CHECK_TEXT_OF_NO_NAME_OCC10, "wnd\[0\]/usr/lbl\[67,8\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_FAGLB03_0120_CHECK_TEXT_OF_NO_NAME_OCC11, "wnd\[0\]/usr/lbl\[67,9\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_FAGLB03_0120_CHECK_TEXT_OF_NO_NAME_OCC12, "wnd\[0\]/usr/lbl\[67,10\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_FAGLB03_0120_CHECK_TEXT_OF_OTH_EXP_NONDED_OCC2, "wnd\[0\]/usr/lbl\[9,7\]")


Call ClickButton("Display Document   \(Shift\+F2\)",False)
Wait 2
Call TakeScreenSHot()
Call ClickButtoniFexist("Call Up Document Overview   \(F9\)",False)
Wait 2
Call TakeScreenSHot()

'''
Call ClickContextButtonToolBar("&MB_VARIANT",1)
Call TakeScreenShot()
Call SelectMenuIdToolBar("&COL0",1)
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot()
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FBL3N_SEARCH_TERM,True)
Call TakeScreenShot()
Call ClickButton("OK   \(Enter\)",True)
Call TakeScreenShot()
Call ClickButton("Cancel   \(F12\)",True)
Call TakeScreenShot()
Call ClickButton("Show Selected Fields \(F7\)",True)
Call ClickButton("Transfer   \(Enter\)",True)
Call TakeScreenShot()
'
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 3, "KTONR", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
Call VerifyGridCellContent("", 4, "KTONR", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KTONR)

Call VerifyGridCellContent("", 1, "KOBEZ", 0, Lcase(DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOBEZ))
Call VerifyGridCellContent("", 2, "ZUONR", 0, Lcase(DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR))

Call VerifyGridCellContent("", 1, "AZBET", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("", 2, "AZBET", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("", 3, "AZBET", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AZBET)
Call VerifyGridCellContent("", 4, "AZBET", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_AZBET)

Call VerifyGridCellContent("", 1, "PSWBT", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PSWBT)
Call VerifyGridCellContent("", 2, "PSWBT", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PSWBT)
Call VerifyGridCellContent("", 3, "PSWBT", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_PSWBT)
Call VerifyGridCellContent("", 4, "PSWBT", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_PSWBT)

Call VerifyGridCellContent("", 1, "Currency", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PSWSL)
Call VerifyGridCellContent("", 2, "Currency", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HWAER)
Call VerifyGridCellContent("", 3, "Currency", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_PSWSL)
Call VerifyGridCellContent("", 4, "Currency", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_PSWSL)


Call ClickButtonifExist("Display Document   \(Shift\+F2\)",False)
Wait 2
Call TakeScreenSHot()

Call ClickButtonifExist("Document Display: General Ledger View   \(Ctrl\+F9\)",FAlse)
Wait 2
Call TakeScreenSHot()
Call ClickButtonifExist("Select Other Ledger   \(Ctrl\+F8\)",FAlse)
Wait 2
Call TakeScreenSHot()

Call VerifyGridCellContent("", 1, "KTONR", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR_OCC1)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR_OCC1)
Call VerifyGridCellContent("", 3, "KTONR", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR_OCC1)
Call VerifyGridCellContent("", 4, "KTONR", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KTONR_OCC1)

Call VerifyGridCellContent("", 1, "KOBEZ", 0, Lcase(DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOBEZ_OCC1))
Call VerifyGridCellContent("", 2, "ZUONR", 0, Lcase(DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_OCC1))

Call VerifyGridCellContent("", 1, "AZBET", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET_OCC1)
Call VerifyGridCellContent("", 2, "AZBET", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET_OCC1)
Call VerifyGridCellContent("", 3, "AZBET", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AZBET_OCC1)
Call VerifyGridCellContent("", 4, "AZBET", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_AZBET_OCC1)

Call VerifyGridCellContent("", 1, "PSWBT", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PSWBT_OCC1)
Call VerifyGridCellContent("", 2, "PSWBT", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PSWBT_OCC1)
Call VerifyGridCellContent("", 3, "PSWBT", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_PSWBT_OCC1)
Call VerifyGridCellContent("", 4, "PSWBT", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_PSWBT_OCC1)

Call VerifyGridCellContent("", 1, "Currency", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PSWSL_OCC1)
Call VerifyGridCellContent("", 2, "Currency", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PSWSL)
Call VerifyGridCellContent("", 3, "Currency", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_HWAER)
Call VerifyGridCellContent("", 4, "Currency", 0, DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_PSWSL_OCC1)
'

Call VerifyTextBoxContent("Document Date","BKPF-BLDAT","",ConvertDate(DT_FAGLB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_ZFBDT_OCC1),False)

Call LogOff()
Call FinalStatus()





