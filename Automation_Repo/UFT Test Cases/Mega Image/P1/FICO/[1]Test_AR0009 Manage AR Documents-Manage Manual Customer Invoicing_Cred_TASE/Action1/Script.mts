
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AR0009 Manage AR Documents-Manage Manual Customer Invoicing_Cred
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


gstrTestCaseName = "Test_AR0009 Manage AR Documents-Manage Manual Customer Invoicing_Cred"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'
'----------------------Tcode FB75----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
Call TakeScreenShot()


'Enter the Company Code
Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_FB75_1000_COMPANY_CODE) 
Call TakeScreenShot()

'Click on Continue
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(2)

Call SetCombo("RF05A-BUSCS","Credit memo")


'Enter the Details
Call SetTextbox("Customer","INVFO-ACCNT","",DT_FB75_0510_CUSTOMER,FALSE)
Call PressEnter()

Call SetTextbox("Document date","INVFO-BLDAT","",DT_FB75_0510_DOCUMENT_DATE,FALSE)
Call SetTextbox("Posting Date","INVFO-BUDAT","",DT_FB75_0510_POSTING_DATE,FALSE)
Call PressEnter()
Call PressEnter()
Call SetTextbox("Reference","INVFO-XBLNR","",DT_FB75_0510_REFERENCE,FALSE)
Call PressEnter()
Call PressEnter()


'Navigate to the Details Tab
Call SelectTab("TS","Details",False)
Wait(1)
Call PressEnter()

'Enter the Header Details
Call SetTextbox("HeaderText","INVFO-BKTXT","",DT_FB75_0550_HEADERTEXT,FALSE)
Call TakeScreenShot()

'Navigate to the Basic data Tab
Call SelectTab("TS","Basic data",False)
Wait(1)
Call PressEnter()

Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB75_0510_AMOUNT,FALSE)
Call SetTextbox("Amount","INVFO-WAERS","",DT_FB75_0510_AMOUNT_OCC1,FALSE)
Call SetTextbox("Tax Amount","INVFO-WMWST",""," ",FALSE)
Call SetTextbox("Text","INVFO-SGTXT","",DT_FB75_0510_TEXT,FALSE)
Call SetTableData("SAPLFSKBTABLE","G/L Acct","1","","",DT_FB75_0100_TABLECELL_GL_ACCT_0,False)
Call SetTableData("SAPLFSKBTABLE","Amount in doc.curr.","1","","",DT_FB75_0100_TABLECELL_AMOUNT_IN_DOCCURR_0,False)
Call SetTableData("SAPLFSKBTABLE","Tax Code","1","","",DT_FB75_0100_TABLECELL_TAX_CODE_0,False)
Call SetTableData("SAPLFSKBTABLE","Assignment","1","","",DT_FB75_0100_TABLECELL_ASSIGNMENT_0,False)
Call SetTableData("SAPLFSKBTABLE","Business Area","1","","",DT_FB75_0100_TABLECELL_BUSINESS_AREA_0,False)
Call SetTableData("SAPLFSKBTABLE","Cost Center","1","","",DT_FB75_0100_TABLECELL_COST_CENTER_0,False)
Call TakeScreenShot()
Call PressEnter()

'Navigate to the Payment Tab
Call SelectTab("TS","Payment",False)
Wait(1)
Call TakeScreenShot()


Call SetTextbox("Bline Date","INVFO-ZFBDT","",ConvertDate(DT_FB75_0520_BLINE_DATE),FALSE)
Call TakeScreenShot()

Call PressEnter()
Call PressEnter()


'Click on Simulate Document Posting
Call ClickButton("Simulate Document Posting   \(F9\)",False)
Wait(5)
Call ClickButton("Simulate Document Posting   \(F9\)",False)
Wait(2)
Call TakeScreenShot()

'''Call ClickLabel("500,00-",0,False)
Call ClickLabel("500.00-",0,False)
Call TakeScreenShot()
'''Call SendKey("{F2}")

'Click More Data
Call ClickButton("Display Additional Data for Document Item   \(F7\)",False) 

Call SelectCheckbox("BSEG-XNEGP",0,"ON",False)
Call TakeScreenShot()

'Click Back
Call ClickButton("Back   \(F3\)",False) 
Call PressEnter()
Call TakeScreenshot()

Call ClickLabel("500.00",0,False)
'''Call SendKey("{F2}")

'Click More Data
Call ClickButton("Display Additional Data for Document Item   \(F7\)",False) 

Call SelectCheckbox("BSEG-XNEGP",0,"ON",False)
Call TakeScreenShot()

'Click Back
Call ClickButton("Back   \(F3\)",False) 

'Click Post
Call ClickButton("Post   \(Ctrl\+S\)",False) 

'Validate If Purchase order is generated
Call GetStatusBar("item1","DT_DOC_NUMBER_OUTPUT")
VerifyStatusBar("Document "&DT_DOC_NUMBER_OUTPUT&" was posted in company code RO02")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",4)
'Click Cancel
'Call ClickButton("Cancel   \(F12\)",True)

Call ClickButtonIfExist("Cancel   \(F12\)",True)

'''''''''''''''''''''----------------------Tcode OB08----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
Call SetTcode(DT_FB75_0100_OKCD_OCC1) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)
Call TakeScreenShot()

'Click Position button
Call ClickButton("VIM_POSI_PUSH",False)  

'Enter details
Call SetTextboxNolabel("SVALD-VALUE","",DT_FB75_0300_EXCH_RATE_TYPE,True)
Call SetTextbox("From currency","SVALD-VALUE","",DT_FB75_0300_FROM_CURRENCY,True)
Call SetTextbox("To-currency","SVALD-VALUE","",DT_FB75_0300_TOCURRENCY,True)
'Call SetTextbox("Valid from","SVALD-VALUE","",ConvertDate(DT_FB75_0300_VALID_FROM),True)
Call SetTextbox("Valid from","SVALD-VALUE","",DT_FB75_0300_VALID_FROM,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
'Set objWindow= SetSAPwindowObj(blnIsItPopup)
'strRowNum = SAPGuisession(sessionObject).sapguiwindow(objWindow).sapguitable(guiTable,"name:=SAPL0SAPTCTRL_V_TCURR").FindRowByCellContent("From","EUR")
'For Iterator = strRowNum To 500000
'	strTO = SAPGuisession(sessionObject).sapguiwindow(objWindow).sapguitable(guiTable,"name:=SAPL0SAPTCTRL_V_TCURR").GetCellData(strRowNum,"To")
'	strFROM = SAPGuisession(sessionObject).sapguiwindow(objWindow).sapguitable(guiTable,"name:=SAPL0SAPTCTRL_V_TCURR").GetCellData(strRowNum,"From")
'	dateValidFrom = SAPGuisession(sessionObject).sapguiwindow(objWindow).sapguitable(guiTable,"name:=SAPL0SAPTCTRL_V_TCURR").GetCellData(Iterator,"ValidFrom")
'	If strTO = "RON" and strFROM = "EUR" and ConvertDate(dateValidFrom) <= ConvertDate(DT_FB75_0510_DOCUMENT_DATE) and ConvertDate(dateValidFrom) > ConvertDate(DT_FB75_0510_DOCUMENT_DATE_OCC1) Then
'	   	'strRowNum1 = SAPGuisession(sessionObject).sapguiwindow(objWindow).sapguitable(guiTable,"name:=SAPL0SAPTCTRL_V_TCURR").FindRowByCellContent("To","RON")
'	   	'For Iterator1 = Iterator To 500000
'	   		'dateValidFrom = SAPGuisession(sessionObject).sapguiwindow(objWindow).sapguitable(guiTable,"name:=SAPL0SAPTCTRL_V_TCURR").GetCellData(Iterator,"ValidFrom")
'	        'If ConvertDate(dateValidFrom) <= ConvertDate(DT_FB75_0510_DOCUMENT_DATE) and ConvertDate(dateValidFrom) > ConvertDate(DT_FB75_0510_DOCUMENT_DATE_OCC1) Then
'	        Call GetTableCellData("SAPL0SAPTCTRL_V_TCURR","Dir.quot.",Iterator,"","","DT_DIR_QUOT_VAL",False)
'        'Exit For
' Exit For 
'	        'End If
'	   ' Iterator1=Iterator1+1    
'	   	'Next
'   End If	
'Iterator=Iterator+1	
'Next
'''Posi_Row = Replace(Replace(Replace(SAPGuiSession("transaction:=OB08").SAPGuiWindow("transaction:=OB08").SAPGuiEdit("name:=VIM_POSITION_INFO").GetROProperty("value"),No_Of_Record,""),"Entry ",""),".","")
'''iRow = Split(Posi_Row," ")
''Call GetTableCellData("SAPL0SAPTCTRL_V_TCURR","Dir.quot.",Posi_Row,"ValidFrom",DT_OB08_0300_VALID_FROM,"DT_OB08_0020_CHECK_TEXT_OF_TABLECELL_DIRQUOT_0_OUTPUT",False)
Call GetTextBoxvalue("VIM_POSITION_INFO",0,"DT_ROW_OUTPUT_2",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'
''''Call GetTableCellData("SAPL0SAPTCTRL_V_TCURR","Dir.quot.",iRow(0),"","","DT_DIR_QUOT_VAL",False)
Call GetTableCellData("SAPL0SAPTCTRL_V_TCURR","Dir.quot.",DT_ROW_EXRT,"","","DT_DIR_QUOT_VAL",False)
'Dim DEC_DIR_QUOT_VAL,objWindow
'Set objWindow= SetSAPwindowObj(blnIsItPopup)
'DEC_DIR_QUOT_VAL=SAPGuisession(sessionObject).sapguiwindow(objWindow).sapguitable(guiTable,"name:=SAPL0SAPTCTRL_V_TCURR").GetCellData(iRow(0),"Dir.quot.")
'Call WriteRunTimeDataToExcel ("DT_DIR_QUOT_VAL",DEC_DIR_QUOT_VAL)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call TakeScreenShot

'Click Back
Call ClickButton("Back   \(F3\)",False) 

Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC5)


'----------------------Tcode FB03----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTcode(DT_FB75_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
Call TakeScreenShot()


Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB75_0100_DOCUMENT_NUMBER_1,FALSE)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB75_0100_COMPANY_CODE,FALSE)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FB75_0100_FISCAL_YEAR,FALSE)
Call TakeScreenShot()
Call PressEnter()

Call VerifyGridCellContent("", 1, "Posting Key", 0, DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "Posting Key", 0, DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

Call DoubleClickGuiGridCell("",0, 1, "Company code", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Acc","BSEG-HKONT",0,DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT, False)
Call VerifyTextBoxContent("Amount in LC","BSEG-DMBTR",0,DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMBTR, False)
Call ClickButton("Back   \(F3\)", False)

Call DoubleClickGuiGridCell("",0,2, "Company code", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Account","BSEG-HKONT",0,DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT, False)
Call VerifyTextBoxContent("Amt\.in loc\.cur\.","BSEG-DMBTR",0,DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMBTR, False)
Call ClickButton("Back   \(F3\)", False)

Call VerifyGridCellContent("", 1, "Account", 0, DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "Account", 0, DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)

Call GetTextboxValue("BKPF-XBLNR", "", "DT_FB75_0750_GET_TEXT_OF_REFERENC_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_FB75_0750_GET_TEXT_OF_REFERENC_OUTPUT",DT_FB75_0750_CHECK_TEXT_OF_REFERENCE)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call VerifyGridCellContent("", 1, "Assignment", 0, DT_FB75_0750_CHECK_TEXT_OF_REFERENCE)
Call VerifyGridCellContent("", 1, "AZBET", 0, DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("", 2, "AZBET", 0, DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("", 1, "Currency", 0, DT_FB75_0750_CHECK_TEXT_OF_CURRENCY)
Call VerifyGridCellContent("",1, "Negative posting", 0, DT_NEG_POSTING)
Call VerifyGridCellContent("",2, "Negative posting", 0, DT_NEG_POSTING)
Call ClickButton("Document Display: General Ledger View   \(Ctrl\+F9\)", False)
Call TakeScreenShot
Call SelectRowGuiGridbyRowNo("",0,1,False)

Call ClickButtonToolBar("&DETAIL",0)
'''Call VerifyGridCellContent("",23, "Cell Content", 0,DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT_OCC1)
'Call VerifyGridCellContent("",17, "Cell Content", 0,DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_OCC1)
Call VerifyGridCellContent("",11, "Cell Content", 0,DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_OCC1)
Call VerifyGridCellContent("",6, "Cell Content", 0,DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT_OCC1)
'Call VerifyGridCellContent("",24, "Cell Content", 0,DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT_OCC1)
Call TakeScreenShot
Call ClickButton("Close window   \(Enter\)", True)
Call SelectRowGuiGridbyRowNo("",0,2,False)

Call ClickButtonToolBar("&DETAIL",0)
'''Call VerifyGridCellContent("",21, "Cell Content", 0,DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT_OCC1)
Call VerifyGridCellContent("",15, "Cell Content", 0,DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT_OCC1)
'Call VerifyGridCellContent("",22, "Cell Content", 0,DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT_OCC1)
Call TakeScreenShot
Call ClickButton("Close window   \(Enter\)", True)

Call VerifyGridCellContent("", 1, "Account", 0, DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR_OCC1)
Call VerifyGridCellContent("", 2, "Account", 0, DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR_OCC1)
Call VerifyGridCellContent("", 1, "Assignment", 0, DT_FB75_0120_CHECK_TEXT_OF_HS1000006561)
Call SelectMenuBar("Environment;Correspondence")
Call TakeScreenShot

'Call ClickLabel("ZRO19", "", True)
Call ClickLabel(DT_FB03_CORRESPONDENCE, "", True)
Call SetTextbox("Document Number","RF022-BELNR","",DT_FB75_0100_DOCUMENT_NUMBER_1,True)
Call SetTextbox("Fiscal Year","RF022-GJAHR","",DT_FB75_1001_FISCAL_YEAR,True)
Call TakeScreenShot

Call ClickButtonIfExist("Copy   \(Enter\)", True)
Call ClickButtonIfExist("Continue   \(Enter\)", True)
'Call VerifyStatusBar("Customer Invoice - MI was requested")
Call VerifyStatusBar(DT_FB03_CORRESPONDENCE_VERIFY)
Call TakeScreenShot

Call ClickButtonIfExist("Back   \(F3\)",False)
Wait(1)
Call ClickButtonIfExist("Back   \(F3\)",False)
Wait(1)

'----------------------Tcode FBL5N----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_FB75_0100_OKCD_OCC2) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC6)
Call TakeScreenShot()

Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_FB75_1000_CUSTOMER_ACCOUNT,False)
'Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_FB75_1000_COMPANY_CODE_OCC1,False)

'Select All Items
Call SelectRadioButton("X_AISEL","All items",False)

'Click on Dynamic selections
Call ClickButton("Dynamic selections   \(Shift\+F4\)", False)

Call ClickButton("%_%%DYN011_%_APP_%-VALU_PUSH", False)


Call SetTableData("SAPLALDBSINGLE","Single value",1,"","",DT_FB75_3010_TABLECELL_SINGLE_VALUE_0,True)

'Click on Copy
Call ClickButton("Copy   \(F8\)", False)

'Call TakeScreenShot()

'Click on Execute
Call ClickButton("Execute   \(F8\)", False)

Call TakeScreenShot()
Call VerifyifGuiLabelExists_ByIndex("S_LEDR",0)
Call VerifyifGuiLabelExists_ByIndex(DT_FB75_0120_CHECK_TEXT_OF_HS1000006561,0)
Call VerifyifGuiLabelExists_ByIndex(DT_FB75_0120_CHECK_TEXT_OF_NO_NAME,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_FB75_0120_CHECK_TEXT_OF_NO_NAME_OCC1,0)
Call ClickButton("Last column   \(Ctrl\+F4\)", False)
'Call VerifyifGuiLabelExists_ByIndex(DT_FB75_0120_CHECK_TEXT_OF_NO_NAME_OCC1,0)
Call ClickButton("Change layout...   \(Ctrl\+F8\)", False)
Call SelectRowGuiTableByRow("SAPLSKBHTC_FIELD_LIST",13, True)
Call ClickButton("Show Sel. Fields \(Ctrl\+F3\)", True)
Call TakeScreenShot
Call ClickButton("Copy   \(Enter\)", True)
Call SetHorizontalScrollBar(30,False)

Call VerifyifGuiLabelExists_ByIndex(DT_FB75_0120_CHECK_TEXT_OF_NO_NAME_OCC2,0)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)", False)
wait 10
Call ClickButton("Back   \(F3\)", False)
wait 10

'''''''''--------TransactionCode-/F.62----------''''

Call SetTcode(DT_FB75_0100_OKCD_OCC3)
''Call SetTcode(DT_FB75_0100_OKCD_OCC3)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC8)

Call SelectCheckbox("NORMBL", 0, "ON", False)
Call SetTextbox("Fiscal year","RGJAHR-LOW","",DT_FB75_1000_FISCAL_YEAR,False)
'Call SetTextbox("Document number","RBELNR-LOW","",DT_FB75_1000_DOCUMENT_NUMBER,False)
Call SetTextbox("Document number","RBELNR-LOW","",DT_FB75_1000_DOCUMENT_NUMBER_OCC1,False)
Call SetTextbox("Company code","RBUKRS-LOW","",DT_FB75_1000_COMPANY_CODE_OCC2,False)
Call SetTextbox("Document type","RBLART-LOW","",DT_FB75_1000_DOCUMENT_TYPE,False)
Call SetTextbox("Correspondence","REVENT","",DT_FB75_1000_CORRESPONDENCE,False)
Call ClickButton("Execute   \(F8\)", False)
Call PressEnter()

'Call ClickButtonIfExist("Yes", True)
Call ClickButtonIfExist("Continue   \(Enter\)", True)
Call ClickButtonIfExist("Yes", True)
'Call ClickButtonIfExist("Continue   \(Enter\)", True)

Call SetTextbox("Output Device","USR01-SPLD","",DT_FB75_1100_OUTPUT_DEVICE,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)", True)
Call ClickButtonIfExist("Yes", True)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)", False)
Call TakeScreenShot

Call VerifyifGuiLabelExists_ByIndex(DT_FB75_0120_CHECK_TEXT_OF_RO02,0)
Call VerifyifGuiLabelExists_ByIndex(DT_FB75_0120_CHECK_TEXT_OF_CUSTOMER_CREDIT_MEMO__MI,0)
Call ClickButton("Back   \(F3\)", False)
Call ClickButton("Back   \(F3\)", False)

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

