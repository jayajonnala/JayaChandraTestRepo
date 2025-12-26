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
'.................Test Script Name : Test_Reverse AP Document_p2_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 12th April
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Reverse AP Document_p2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Reverse AP Document_p2_TASE.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

''----------------------------------------------F-02--------------------------------------
Call SetTcode(DT_TRANSACTIONCODE_TAX_PRE) 
Call PressEnter() 
Call CheckTCodeScreen(DT_TRANSACTIONCODE_TAX_PRE)
Call TakeScreenShot()


'Call SelectMenuBar("Settings;Processing Options")
Call SelectMenuBar("Settings;Editing Options")
Call TakeScreenShot()
Call SelectCheckbox("RFOPT2-XSNET",8,DT_F_02_CHECK_BOX_STATUS_PRE,false)
Call TakeScreenShot()

Call ClickButton("Change user master   \(Ctrl\+S\)",False)
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot()
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot()

Call ClickButtonIfExist("Yes",True)
Call TakeScreenShot()

'---------------------------------------------------------------------------------------------------------------


Call WriteRunTimeDataToExcelGlobalSheet ("DT_REF_INC",(Cint(DT_REF_INC)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'''----------------------Tcode FB65----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

'this step is not in the log but it is a mandate in screen
Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_company_code)
'''Call SetTextbox("Company Code","BKPF-BUKRS","",DT_company_code,True)
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot()

'Call SetTextbox("Vendor","INVFO-ACCNT","",DT_FB65_0010_VENDOR,False)
Call SetTextboxNoLabel("INVFO-ACCNT","",DT_FB65_0010_VENDOR,False)
Call SetTextbox("Document date","INVFO-BLDAT","",Replace((DT_FB65_0010_DOCUMENT_DATE),"/","."),False)
Call SetTextbox("Reference","INVFO-XBLNR","",DT_FB65_0010_REFERENCE,False)
'Call SetTextbox("Posting Date","INVFO-BUDAT","",Replace((DT_FB60_0010_POSTING_DATE),"/","."),False)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB65_0010_AMOUNT,False)
Call SetTextbox("Incg Doc\. Nmbr","INVFO-INWARDNO_HD","",DT_FB60_0010_REFERENCE,False)
Call SelectCheckbox("INVFO-XMWST",0,DT_FB65_0010_CALCULATE_TAX,False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()

Call SetTableDataNoRef("SAPLFSKBTABLE","G/L acct",1,DT_FB65_0100_TABLECELL_GL_ACCT_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","D/C",1,DT_FB65_0100_TABLECELL_DC_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Amount in doc.curr.",1,DT_FB65_0100_TABLECELL_AMOUNT_IN_DOCCURR_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Tax code",1,DT_FB65_0100_TABLECELL_TAX_CODE_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Company code",1,DT_FB65_0100_TABLECELL_COMPANY_CODE_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Business area",1,DT_FB65_0100_TABLECELL_BUSINESS_AREA_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Cost center",1,DT_FB65_0100_TABLECELL_COST_CENTER_0,False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()

Call SelectTab("TS","Payment",False)
Call TakeScreenShot()

Call SetTextbox("BaselineDt","INVFO-ZFBDT","",Replace((DT_FB65_0020_BASELINEDT),"/","."),False)
Call PressEnter()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call PressEnter()
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
Call GetStatusBar("item1","DT_FB65_1100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetStatusBar("item2","DT_FB65_1100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_FB65_1100_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot()
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Wait(1)


''''----------------------Tcode F-44----------------------------

Call SetTcode(DT_FB65_1100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_FB65_1100_OKCD)
Call TakeScreenShot()

Call SelectRadioButton("RF05A-XPOS1",DT_FB65_0131_DOCUMENT_NUMBER,False)
Call SetTextbox("Account","RF05A-AGKON","",DT_FB65_0131_ACCOUNT,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB65_0100_TABLECELL_COMPANY_CODE_0,False)
'Call SetTextbox("Special G/L Ind","RF05A-AGUMS","",DT_F44_0131_SPECIAL_GL_IND,False)
'Call SetTextbox("Clearing Date","BKPF-BUDAT","",Replace(DT_F44_0131_CLEARING_DATE,"/","."),False)
Call TakeScreenShot()

Call PressEnter() 
Wait(1)
Call TakeScreenShot()

Call SetTextbox("From","RF05A-SEL01",0,DT_FB65_0731_FROM1,False)
Call SetTextbox("From","RF05A-SEL01",1,DT_FB65_0731_FROM_OCC1,False)

Call TakeScreenShot()

Call PressEnter() 
Wait(1)
Call TakeScreenShot()

Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)
'
Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
Wait(1)
Call TakeScreenShot()

Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC2)

Call ClickButton("Select All",False) 
Call TakeScreenShot()

Call ClickButton("Deactivate Items",False)
Call TakeScreenShot()

Call ClickButton("Field content search",False) 
Call TakeScreenShot()

Call SelectRadioButton("RF05A-XPOS1","Document Number",True)
Call TakeScreenShot()

Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot()

Call SetTextbox("From","RF05A-SEL01","",DT_FB65_0731_FROM1,True)
Call SetTextbox("From","RF05A-SEL01",1,DT_FB65_0731_FROM_OCC1,True)'''''''
Call TakeScreenShot()

Call PressEnter() 
Call TakeScreenShot()

Call ClickButton("Select All",False) 
Call TakeScreenShot()

Call ClickButton("Activate Items",False)
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
wait(1)
Call TakeScreenShot()
Call GetStatusBar("item1","DT_FB65_0131_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetStatusBar("item2","DT_FB65_0131_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_FB65_0131_CHECK_TEXT_OF_STATUSBAR)
Call VerifyStatusBarMessageType("S")

Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_company_code)
Call ClickButtonIfExist("Continue   \(Enter\)",True)



'----------------------------------F-02---------------------------------------------
Call SetTcode(DT_TRANSACTIONCODE_TAX_POST) 
Call PressEnter() 
Call CheckTCodeScreen(DT_TRANSACTIONCODE_TAX)
Call TakeScreenShot()

'Call SelectMenuBar("Settings;Processing Options")
Call SelectMenuBar("Settings;Editing Options")
Call TakeScreenShot()
Call SelectCheckbox("RFOPT2-XSNET",8,DT_F_02_CHECK_BOX_STATUS_POST,false)
Call TakeScreenShot()

Call ClickButton("Change user master   \(Ctrl\+S\)",False)
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot()
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot()

Call ClickButtonIfExist("Yes",True)
Call TakeScreenShot()

'-----------------------------------------------------------------------------------------


Call LogOff()
Call FinalStatus ()
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

'Call SelectTab("TS","Details",False)
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
'Call SetTextbox("BaselineDt","INVFO-ZFBDT","",Replace((DT_FB65_0020_BASELINEDT),"/","."),False)


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
'
'Call ClickButton("Simulate Document Posting   \(F9\)",False)
'Wait(1)
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

'Call PressEnter()
''Capture the screenshot
'Call TakeScreenShot()

''validate components
'Call VerifyifGuiLabelExists(DT_FB65_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
'Call VerifyifGuiLabelExists(DT_FB65_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
'Call VerifyifGuiLabelExists(DT_FB65_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
''
'''Capture the screenshot
''Call TakeScreenShot()
'''Click on Post Buton
''Call ClickButton("Post   \(Ctrl\+S\)",False)
''Call PressEnter()
''Call ClickButtonIfExist("Save",True)
''wait(1)
''Call TakeScreenShot()
'''Validate If document is posted and get the status bar nummber
''Call GetStatusBar("item1","DT_FB65_1100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
''Call GetStatusBar("item2","DT_FB65_1100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''
'''Call VerifyStatusBarMessageType("S")
''VerifyStatusBar(DT_FB65_1100_CHECK_TEXT_OF_STATUSBAR)
'''Capture the screenshot
''Call TakeScreenShot()
''Call ClickButton("Cancel   \(F12\)",True)
''Wait(1)
''
'''
'''''----------------------Tcode F-44----------------------------
'''
''''Enter the Tcode
''Call SetTcode(DT_FB65_1100_OKCD) 
''Call PressEnter()     ' 
''Call CheckTCodeScreen(DT_FB65_1100_OKCD)
'''Capture the screenshot
''Call TakeScreenShot()
''
''Call SelectRadioButton("RF05A-XPOS1",DT_FB65_0131_DOCUMENT_NUMBER,False)
''Call SetTextbox("Account","RF05A-AGKON","",DT_FB65_0131_ACCOUNT,False)
''Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB65_0100_TABLECELL_COMPANY_CODE_0,False)
'''Call SetTextbox("Special G/L Ind","RF05A-AGUMS","",DT_F44_0131_SPECIAL_GL_IND,False)
'''Call SetTextbox("Clearing Date","BKPF-BUDAT","",Replace(DT_F44_0131_CLEARING_DATE,"/","."),False)
''
'''Capture the screenshot
''Call TakeScreenShot()
''
''Call PressEnter() 
'''Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
''
''Wait(1)
'''Capture the screenshot
''Call TakeScreenShot()
''
'''set filter criteria
''Call SetTextbox("From","RF05A-SEL01",0,DT_FB65_0731_FROM,False)
''Call SetTextbox("From","RF05A-SEL01",1,DT_FB65_0731_FROM_OCC1,False)
''
'''Capture the screenshot
''Call TakeScreenShot()
''
''Call PressEnter() 
'''Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
''Wait(1)
'''Capture the screenshot
''Call TakeScreenShot()
''
''Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)
'''
'''Call ClickButton("Back   \(F3\)",False)
''''Capture the screenshot
'''Call TakeScreenShot()
'''
''Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
''Wait(1)
'''Capture the screenshot
''Call TakeScreenShot()
''
''Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC2)
''
''Call ClickButton("Select All",False) 
'''Capture the screenshot
''Call TakeScreenShot()
''
''Call ClickButton("Deactivate Items",False)
'''Capture the screenshot
''Call TakeScreenShot()
''
''Call ClickButton("Field content search",False) 
''Call TakeScreenShot()
''
''Call SelectRadioButton("RF05A-XPOS1","Document Number",True)
''Call TakeScreenShot()
''
''Call ClickButton("Continue   \(Enter\)",True)
''Call TakeScreenShot()
''
''Call SetTextbox("From","RF05A-SEL01","",DT_FB65_0731_FROM,True)
''Call SetTextbox("From","RF05A-SEL01",1,DT_FB65_0731_FROM_OCC1,True)'''''''
''Call TakeScreenShot()
''
''Call PressEnter() 
''Call TakeScreenShot()
''
''Call ClickButton("Select All",False) 
'''Capture the screenshot
''Call TakeScreenShot()
''
''Call ClickButton("Activate Items",False)
'''Capture the screenshot
''Call TakeScreenShot()
''
'''Call ClickButton("Document Overview   \(Shift\+F2\)",False)
''''Capture the screenshot
'''Call TakeScreenShot()
''
''Call ClickButton("Post   \(Ctrl\+S\)",False)
''wait(1)
'''Capture the screenshot
''Call TakeScreenShot()
'''Validate If doc is generated
''Call GetStatusBar("item1","DT_FB65_0131_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
''Call GetStatusBar("item2","DT_FB65_0131_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''VerifyStatusBar(DT_FB65_0131_CHECK_TEXT_OF_STATUSBAR)
''Call VerifyStatusBarMessageType("S")
''
''
'''Log Off From Applicaton
''Call LogOff()
''Call FinalStatus ()

