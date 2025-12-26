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


'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Clear AP Accounts - Manual_p2_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 7th April
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Clear AP Accounts - Manual_p2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Clear AP Accounts - Manual_p2_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'Increment the parameter
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'''----------------------Tcode FB65----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

'this step is not in the log but it is a mandate in screen
'Call SetTextbox("Company Code","BKPF-BUKRS","",DT_company_code,True)
Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_Company_code)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Continue   \(Enter\)",True)
'Capture the screenshot
Call TakeScreenShot()

'Call SetTextbox("Vendor","INVFO-ACCNT","",DT_FB65_0010_VENDOR,False)
Call SetTextboxNoLabel("INVFO-ACCNT","",DT_FB65_0010_VENDOR,False)
Call SetTextbox("Document date","INVFO-BLDAT","",Replace((DT_FB65_0010_DOCUMENT_DATE),"/","."),False)
Call SetTextbox("Reference","INVFO-XBLNR","",DT_FB65_0010_REFERENCE,False)
Call SetTextbox("Incg Doc. Nmbr","INVFO-INWARDNO_HD","",DT_FB65_0010_REFERENCE,False)
'Call PressEnter()
'Call SetTextbox("Posting Date","INVFO-BUDAT","",Replace((DT_FB60_0010_POSTING_DATE),"/","."),False)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB65_0010_AMOUNT,False)
'Select Calculate Tax field as true
Call SelectCheckbox("INVFO-XMWST",0,DT_FB65_0010_CALCULATE_TAX,False)

''Capture the screenshot
'Call TakeScreenShot()
'Call PressEnter()

'Capture the screenshot
Call TakeScreenShot()

'set table data
Call SetTableDataNoRef("SAPLFSKBTABLE","G/L acct",1,DT_FB65_0100_TABLECELL_GL_ACCT_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Amount in doc.curr.",1,DT_FB65_0100_TABLECELL_AMOUNT_IN_DOCCURR_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Tax code",1,DT_FB65_0100_TABLECELL_TAX_CODE_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Business area",1,DT_FB65_0100_TABLECELL_BUSINESS_AREA_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Cost center",1,DT_FB65_0100_TABLECELL_COST_CENTER_0,False)
'Capture the screenshot
Call TakeScreenShot()
'
'Call PressEnter()
''Capture the screenshot
'Call TakeScreenShot()
'
'Call ClickButton("Enter",False)
''Capture the screenshot
'Call TakeScreenShot()
'
'Call ClickButton("Enter",False)
''Capture the screenshot
'Call TakeScreenShot()
'
'Call SelectTab("TS","Payment",False)
''Capture the screenshot
'Call TakeScreenShot()
'
'Call ClickButton("Enter",False)
''Capture the screenshot
'Call TakeScreenShot()

Call SelectTab("TS","Details",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Enter",False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TS","Payment",False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("BaselineDt","INVFO-ZFBDT","",Replace((DT_FB65_0020_BASELINEDT),"/","."),False)


'Call SelectTab("TS","Tax",False)
''Capture the screenshot
'Call TakeScreenShot()
'
'Call ClickButton("Enter",False)
''Capture the screenshot
'Call TakeScreenShot()
'
'Call SelectTab("TS","Amount split",False)
''Capture the screenshot
'Call TakeScreenShot()
'
'Call ClickButton("Enter",False)
''Capture the screenshot
'Call TakeScreenShot()
'
'Call ClickButtonIfExist("Continue   \(Enter\)",True)
''Capture the screenshot
'Call TakeScreenShot()

Call ClickButton("Simulate Document Posting   \(F9\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Enter",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Enter",False)
'Capture the screenshot
Call TakeScreenShot()

'Call PressEnter()
''Capture the screenshot
'Call TakeScreenShot()

'validate components
Call VerifyifGuiLabelExists(DT_FB65_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyifGuiLabelExists(DT_FB65_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyifGuiLabelExists(DT_FB65_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)


'Capture the screenshot
Call TakeScreenShot()
'Click on Post Buton
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
'Validate If document is posted and get the status bar nummber
Call GetStatusBar("item1","DT_FB65_1100_DOCUMENT_NUMBER_OUTPUT")
'Call GetStatusBar("item2","DT_FB60_1100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call VerifyStatusBarMessageType("S")
VerifyStatusBar(DT_FB65_1100_CHECK_TEXT_OF_STATUSBAR)
'Capture the screenshot
Call TakeScreenShot()


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

