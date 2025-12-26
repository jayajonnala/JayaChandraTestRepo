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
'.................Test Script Name : Test_Reverse AP Document_p9_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 14th April
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Reverse AP Document_p9_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Reverse AP Document_p9_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_REFERENCE",(Cint(DT_INCREMENT_REFERENCE)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'''----------------------Tcode FB08----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

Call SetTextbox("Document Number","RF05A-BELNS","",DT_FB08_0105_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB08_0105_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05A-GJAHS","",DT_FB08_0105_FISCAL_YEAR,False)
Call SetTextbox("Reversal Reason","UF05A-STGRD","",DT_FB08_0105_REVERSAL_REASON,False)
Call TakeScreenShot()

'Call PressEnter() 
''Capture the screenshot
'Call TakeScreenShot()

'Call ClickButton("Display document before reversal   \(F5\)",False)
''Capture the screenshot
'Call TakeScreenShot()
'
''validate grid components
'call VerifyGridCellContent("",1,"Description","",Left(DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOBEZ,4))
'call VerifyGridCellContent("",2,"Description","",DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KOBEZ)
'
''Capture the screenshot
'Call TakeScreenShot()
'
'Call ClickButton("Back   \(F3\)",False)
''Capture the screenshot
'Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
Call GetStatusBar("item1","DT_FB08_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetStatusBar("item2","DT_FB08_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_FB08_0105_CHECK_TEXT_OF_STATUSBAR)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)

Call SelectMenuBar("Document;Display")
Wait(1)
Call TakeScreenShot()

'Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB08_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR,False)
'Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB08_0105_COMPANY_CODE,False)
'Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FB08_0105_FISCAL_YEAR,False)
'Call TakeScreenShot()

'Call PressEnter()
'Call TakeScreenShot()

'Call ClickButton("Display Document Header   \(F5\)",False)
'Call VerifyTextBoxContent("Document type","T003T-LTEXT","",Lcase(DT_FB08_1710_CHECK_TEXT_OF_DOCUMENT_TYPE),True)
'Call TakeScreenShot()

'Call PressEnter()
'Call TakeScreenShot()

call VerifyGridCellContent("",1,"Posting Key","",DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
call VerifyGridCellContent("",2,"Posting Key","",DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
'call VerifyGridCellContent("",3,"Posting Key","",DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)

call VerifyGridCellContent("",1,"Account","",DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
call VerifyGridCellContent("",2,"Account","",DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)

'call VerifyGridCellContent("",2,"Alternative Account No.","",DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT)
'call VerifyGridCellContent("",3,"Alternative Account No.","",DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_LOKKT)

'''----------------------Tcode FB60----------------------------

Call SetTcode(DT_FB08_0750_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_FB08_0750_OKCD)
Call TakeScreenShot()

'this step is not in the log but it is a mandate in screen
Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_FB08_0105_COMPANY_CODE)
''Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB08_0105_COMPANY_CODE,True)
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot()

'Call SetComboByKey("RF05A-BUSCS",DT_FB08_1100_TRANSACTN)
'Call SetTextbox("Vendor","INVFO-ACCNT","",DT_FB08_0010_VENDOR,False)
Call SetTextboxNoLabel("INVFO-ACCNT","",DT_FB08_0010_VENDOR,False)'
Call SetTextbox("Invoice date","INVFO-BLDAT","",Replace((DT_FB08_0010_INVOICE_DATE),"/","."),False)
Call SetTextbox("Reference","INVFO-XBLNR","",DT_FB08_0010_REFERENCE,False)
Call SetTextbox("Posting Date","INVFO-BUDAT","",Replace((DT_FB08_0010_POSTING_DATE),"/","."),False)
Call SetTextbox("Incg Doc\. Nmbr","INVFO-INWARDNO_HD","",DT_FB60_0010_REFERENCE,False)
'Call SetTextbox("Cross-CCode No.","INVFO-BVORG","",DT_FB60_0010_CROSSCC_NO,False)
'Call SetComboByKey("INVFO-BLART",DT_FB08_0010_DOCUMENT_TYPE)

'Call TakeScreenShot()
'Call PressEnter()

Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB08_0010_AMOUNT,False)
Call SetTextbox("Amount","INVFO-WAERS","",DT_FB08_0010_AMOUNT_OCC1,False)
'Select Calculate Tax field as true
Call SelectCheckbox("INVFO-XMWST",0,DT_FB08_0010_CALCULATE_TAX,False)

Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()

'Call SelectTab("TS","Payment",False)
'wait(1)
'Call Clickbuttonifexist("Continue   \(Enter\)",True)
'wait(1)
''Call SetTextbox("BaselineDt","INVFO-ZFBDT","",Replace((DT_FB08_0020_BASELINEDT),"/","."),False)
'Call SetTextbox("Payment Ref.","INVFO-KIDNO","",DT_payment_reference,False)
'''Capture the screenshot
'Call TakeScreenShot()

Call SetTableDataNoRef("SAPLFSKBTABLE","G/L acct",1,DT_FB08_0100_TABLECELL_GL_ACCT_0,False)
'Call SetTableDataNoRef("SAPLFSKBTABLE","D/C",1,DT_FB08_0100_TABLECELL_DC_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Amount in doc.curr.",1,DT_FB08_0100_TABLECELL_AMOUNT_IN_DOCCURR_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Tax code",1,DT_FB08_0100_TABLECELL_TAX_CODE_0,False)
'Call SetTableDataNoRef("SAPLFSKBTABLE","Company code",1,DT_FB08_0100_TABLECELL_COMPANY_CODE_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Business area",1,DT_FB08_0100_TABLECELL_BUSINESS_AREA_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Cost center",1,DT_FB08_0100_TABLECELL_COST_CENTER_0,False)
Call TakeScreenShot()

'Call ClickButton("Enter",False)
''Capture the screenshot
'Call TakeScreenShot()
'
'Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)
'
'Call ClickButton("Enter",False)
''Capture the screenshot
'Call TakeScreenShot()
'
'Call SelectTab("TS","Payment",False)
''Capture the screenshot
'Call TakeScreenShot()
'
'Call PressEnter()
''Capture the screenshot
'Call TakeScreenShot()

'
'Call ClickButton("Enter",False)
''Capture the screenshot
'Call TakeScreenShot()
'
'Call SelectTab("TS","Details",False)
''Capture the screenshot
'Call TakeScreenShot()
'
'Call ClickButton("Enter",False)
''Capture the screenshot
'Call TakeScreenShot()
'
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
'Call PressEnter()
''Capture the screenshot
'Call TakeScreenShot()
'
''validate components
'Call VerifyifGuiLabelExists(DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
'Call VerifyifGuiLabelExists(DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
'Call VerifyifGuiLabelExists(DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
'
'
Call TakeScreenShot()
Call PressEnter()
Call PressEnter()
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
Call GetStatusBar("item1","DT_FB08_1100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetStatusBar("item2","DT_FB08_1100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC2)
VerifyStatusBar(DT_FB08_1100_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot()

Call LogOff()
Call FinalStatus ()

