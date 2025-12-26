'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)


'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Manage Manual Vendor Invoicing Crediting (Non PO Invoices)_p10_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 29th April
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Manage Manual Vendor Invoicing Crediting (Non PO Invoices)_p10_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Manage Manual Vendor Invoicing  Crediting (Non PO Invoices)_p10_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
'''----------------------Tcode FB60----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_Company_code,True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Continue   \(Enter\)",True)
'Call SetTextbox("Vendor","INVFO-ACCNT","",DT_FB60_0010_VENDOR,False)
Call SetTextboxNoLabel("INVFO-ACCNT","",DT_FB60_0010_VENDOR,False)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB60_0010_AMOUNT,False)
Call SetTextbox("Invoice date","INVFO-BLDAT","",Replace((DT_FB60_0010_INVOICE_DATE),"/","."),False)
Call SetTextbox("Reference","INVFO-XBLNR","",DT_FB60_0010_REFERENCE,False)
Call PressEnter()
wait(1)
Call Clickbuttonifexist("Continue   \(Enter\)",True)
Call SetTextbox("Incg Doc. Nmbr","INVFO-INWARDNO_HD","",DT_FB60_0010_REFERENCE,False)
'Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

'Select Calculate Tax field as true
Call SelectCheckbox("INVFO-XMWST",0,DT_FB60_0010_CALCULATE_TAX,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()
Call PressEnter()
Call PressEnter()
'set table data
Call SetTableDataNoRef("SAPLFSKBTABLE","G/L acct",1,DT_FB60_0100_TABLECELL_GL_ACCT_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Amount in doc.curr.",1,DT_FB60_0100_TABLECELL_AMOUNT_IN_DOCCURR_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Tax code",1,DT_TAX_CODE,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Business area",1,DT_FB60_0100_TABLECELL_BUSINESS_AREA_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Cost center",1,DT_FB60_0100_TABLECELL_COST_CENTER_0,False)
Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()
Call Clickbuttonifexist("Continue   \(Enter\)",True)
Call SelectTab("TS",DT_FB60_1100_PAYMENT,False)
wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call Clickbuttonifexist("Continue   \(Enter\)",True)
wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Part\. bank","INVFO-BVTYP","",DT_FB60_0020_PART_BANK,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TS",DT_FB60_1100_DETAILS,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TS",DT_FB60_1100_TAX,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TS",DT_FB60_1100_AMOUNT_SPLIT,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TS",DT_FB60_1100_NOTES,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Enter",False)
Call ClickButton("Enter",False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextArea(DT_FB60_0040_TEXTEDIT_SHELL)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Simulate Document Posting   \(F9\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()

'validate components
Call VerifyifGuiLabelExists(DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyifGuiLabelExists(DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KOBEZ)

'Capture the screenshot
Call TakeScreenShot()
'Click on Post Buton
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
'Validate If document is posted and get the status bar nummber
Call GetStatusBar("item1","DT_FB60_1100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBarMessageType("S")
VerifyStatusBar(DT_FB60_1100_CHECK_TEXT_OF_STATUSBAR)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Cancel   \(F12\)",True)
Wait(1)

'
'''----------------------Tcode FS00----------------------------
'Enter the Tcode
Call SetTcode(DT_FB60_1100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_FB60_1100_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("G/L Account","GLACCOUNT_SCREEN_KEY-SAKNR","",DT_FB60_2011_GL_ACCOUNT,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()
Call VerifyTextBoxContent("Short Text","GLACCOUNT_SCREEN_COA-TXT20_ML","",Lcase(DT_FB60_2101_CHECK_TEXT_OF_SHORT_TEXT),False)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

