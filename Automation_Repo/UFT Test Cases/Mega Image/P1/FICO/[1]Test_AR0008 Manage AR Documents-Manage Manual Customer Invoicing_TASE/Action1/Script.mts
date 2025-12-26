
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AR0008 Manage AR Documents-Manage Manual Customer Invoicing
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

gstrTestCaseName = "Test_AR0008 Manage AR Documents-Manage Manual Customer Invoicing"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


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
Call SetTextbox("Document date","INVFO-BLDAT","",ConvertDate(DT_FB75_0510_DOCUMENT_DATE),FALSE)
Call SetTextbox("Posting Date","INVFO-BUDAT","",ConvertDate(DT_FB75_0510_POSTING_DATE),FALSE)
Call SetTextbox("Reference","INVFO-XBLNR","",DT_FB75_0510_REFERENCE,FALSE)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB75_0510_AMOUNT,FALSE)
Call SetTableData("SAPLFSKBTABLE","G/L Acct","1","","",DT_FB75_0100_TABLECELL_GL_ACCT_0,False)
Call SetTableData("SAPLFSKBTABLE","Amount in doc.curr.","1","","",DT_FB75_0100_TABLECELL_AMOUNT_IN_DOCCURR_0,False)
Call TakeScreenShot()
Call SetTableData("SAPLFSKBTABLE","Tax Code","1","","",DT_FB75_0100_TABLECELL_TAX_CODE_0,False)
Call SetTableData("SAPLFSKBTABLE","Business Area","1","","",DT_FB75_0100_TABLECELL_BUSINESS_AREA_0,False)
Call SetTableData("SAPLFSKBTABLE","Cost Center","1","","",DT_FB75_0100_TABLECELL_COST_CENTER_0,False)
Call SelectCheckbox("INVFO-XMWST",0,"ON",False)
Call TakeScreenShot()
Call PressEnter()


'Navigate to the Details Tab
Call SelectTab("TS","Details",False)
Wait(1)
Call SetTextbox("HeaderText","INVFO-BKTXT","",DT_FB75_0550_HEADERTEXT,FALSE)
Call PressEnter()
Call SelectTab("TS","Details",False)
Wait(1)
Call SelectCheckbox("INVFO-XNEGP",0,"ON",False)
Call TakeScreenShot()

'Navigate to the Payment Tab
Call SelectTab("TS","Payment",False)
Wait(1)

Call SetTextbox("Bline Date","INVFO-ZFBDT","",ConvertDate(DT_FB75_0520_BLINE_DATE),FALSE)
Call TakeScreenShot()

'Navigate to the Basic data Tab
Call SelectTab("TS","Basic data",False)
Wait(1)
Call PressEnter()

'Click on Simulate Document Posting
Call ClickButton("Simulate Document Posting   \(F9\)",False)
Wait(2)
Call TakescreenShot()
Call PressEnter()
Call TakescreenShot()

'Click Post
Call ClickButton("Post   \(Ctrl\+S\)",False) 
''Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False) 

'Validate If Purchase order is generated
Call GetStatusBar("item1","DT_DOC_NUMBER_OUTPUT")
VerifyStatusBar("Document "&DT_DOC_NUMBER_OUTPUT&" was posted in company code RO02")

'Click Cancel
Call ClickButtonIfExist("Cancel   \(F12\)",True) 

'''Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)

'----------------------Tcode FB03----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTcode(DT_FB75_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
Call TakeScreenShot()


Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB75_0100_DOCUMENT_NUMBER,FALSE)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB75_0100_COMPANY_CODE,FALSE)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FB75_0100_FISCAL_YEAR,FALSE)
Call TakeScreenShot()
Call PressEnter()


Call ClickButtonToolBar("&FIND",0)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Negative Posting",True)
Call TakeScreenShot()

'Click OK   \(Enter\)
Call ClickButton("OK   \(Enter\)",True) 

'Click Cancel
Call ClickButton("Cancel   \(F12\)",True) 

Call VerifyGridCellContent("",1,"Negative posting",0,"X")
Call VerifyGridCellContent("",2,"Negative posting",0,"X")


Call GetTextboxValue("BKPF-XBLNR",0,"DT_REFERENCE_NO",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)


'Verify the Grid Content
Call VerifyTextBoxContent("Currency","BKPF-WAERS",0,DT_FB75_0750_CHECK_TEXT_OF_CURRENCY,False)
Call VerifyGridCellContent("",1,"Assignment",0,DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)
Call VerifyGridCellContent("",1,"Posting Key",0,DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",2,"Posting Key",0,DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("",1,"Account",0,DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("",2,"Account",0,DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("",1,"Alternative Account No.",0,DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOKKT)
Call VerifyGridCellContent("",2,"Alternative Account No.",0,DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT)
Call VerifyGridCellContent("",1,"Amount",0,DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("",2,"Amount",0,DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("",1,"Tax Code",0,DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ)
Call VerifyGridCellContent("",2,"Tax Code",0,DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MWSKZ)


'Click on Display Document Header
Call ClickButtonIfExist("Display Document Header   \(F5\)",False)
Wait(2)


Call VerifyTextBoxContent("Document type","BKPF-BLART",0,DT_FB75_1710_CHECK_TEXT_OF_DOCUMENT_TYPE,True)

'Click Continue/Confirm
Call ClickButton("Continue/Confirm   \(Enter\)",True) 

'Navigate to Environment;Correspondence
Call SelectMenuBar("Environment;Correspondence")
Call TakeScreenShot()

'Click Find
Call ClickButton("Find   \(Ctrl\+F\)",True) 

Call SetTextbox("Find","RSYSF-STRING","",DT_FB75_0800_FIND,True)
Call TakeScreenShot()

'Click Find   \(Enter\)
Call ClickButton("Find   \(Enter\)",True) 
Call TakeScreenShot()

Call ClickLabel("Customer credit memo - MI",0,True)

'Click Copy   \(Enter\)
Call ClickButton("Copy   \(Enter\)",True) 


Call SetTextbox("Document Number","RF022-BELNR","",DT_FB75_1001_DOCUMENT_NUMBER,True)
Call SetTextbox("Fiscal Year","RF022-GJAHR","",DT_FB75_1001_FISCAL_YEAR,True)
Call TakeScreenShot()

'Click Continue
Call ClickButton("Continue   \(Enter\)",True) 

Call VerifyStatusBar(DT_FB75_0750_CHECK_TEXT_OF_STATUSBAR)

Call ClickButtonIfExist("Back   \(F3\)",False)
Wait(1)
Call ClickButtonIfExist("Back   \(F3\)",False)
Wait(1)

Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)

'----------------------Tcode F.62----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_FB75_0100_OKCD_OCC1) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)
Call TakeScreenShot()

Call SetTextbox("Company code","RBUKRS-LOW","",DT_FB75_1000_COMPANY_CODE_OCC1,False)
Call SetTextbox("Fiscal year","RGJAHR-LOW","",DT_FB75_1000_FISCAL_YEAR,False)
Call SetTextbox("Document number","RBELNR-LOW","",DT_FB75_1000_DOCUMENT_NUMBER,False)
Call SetTextbox("Document type","RBLART-LOW","",DT_FB75_1000_DOCUMENT_TYPE,False)
Call SetTextbox("Correspondence","REVENT","",DT_FB75_1000_CORRESPONDENCE,False)
Call SelectCheckbox("NORMBL", 0, "ON", False)
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)", False)
Call TakeScreenShot()
'Call SelectCheckbox("NORMBL", 0, "ON", False)

Call PressEnter() 
Call ClickButtonIfExist("Yes", True)
Call SetTextbox("Output Device","USR01-SPLD","",DT_FB75_1100_OUTPUT_DEVICE,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)", True)
Call ClickButtonIfExist("Yes", True)
Call TakeScreenShot
Call VerifyifGuiLabelExists_ByIndex(Lcase(DT_FB75_0120_CHECK_TEXT_OF_CUSTOMER_CREDIT_MEMO__MI),0)
Call ClickButtonIfExist("Back   \(F3\)",False)
Wait(1)


'----------------------Tcode FBL5N----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_SAPTRANSACTIONCODE2) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE2)
Call TakeScreenShot()

Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_FBL5N_1000_CUSTOMER_ACCOUNT,False)
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_FBL5N_1000_COMPANY_CODE,False)

'Select All Items
Call SelectRadioButton("X_AISEL","All items",False)

'Click on Dynamic selections
Call ClickButton("Dynamic selections   \(Shift\+F4\)", False)

Call ClickButton("%_%%DYN011_%_APP_%-VALU_PUSH", False)


Call SetTableData("SAPLALDBSINGLE","Single value",1,"","",DT_FBL5N_3010_TABLECELL_SINGLE_VALUE_0,True)

'Click on Copy
Call ClickButton("Copy   \(F8\)", False)

'Call TakeScreenShot()

'Click on Execute
Call ClickButton("Execute   \(F8\)", False)
Call VerifyifGuiLabelExists_ByIndex(ConvertDate(DT_NET_DUE_DATE),0)
Call TakeScreenShot()


'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************



