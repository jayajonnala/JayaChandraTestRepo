
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_AR0034 Clear AR Account- automatic clearing RON_TASE
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


gstrTestCaseName = "Test_AR0034 Clear AR Account- automatic clearing RON_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\FICO\TASE_DT_AR0001 Manage AR Documents-Manage Manual Customer Invoicing_Cred.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


''''----------------------Tcode F-22----------------------------

'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F22_0100_DOCUMENT_DATE),False)
'Call SetTextbox("Posting Date","BKPF-BUDAT","",Replace((DT_F22_100_POSTING_DATE),"/","."),False)
'Call SetTextbox("Type","BKPF-BLART","",DT_OB08_0100_TYPE,False)
Call SetTextbox("Period","BKPF-MONAT","",DT_F22_0100_PERIOD,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F22_0100_COMPANY_CODE,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F22_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F22_0100_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F22_0100_DOCHEADER_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F22_0100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F22_0100_ACCOUNT,False)
Call SetTextbox("SGL Ind","RF05A-NEWUM","",DT_F22_0100_SGL_IND,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-22").SapGuiWindow("transaction:=F-22").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next

Call SelectCheckbox("BKPF-XMWST","1",DT_F22_0303_CALCULATE_TAX,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F22_0303_AMOUNT,False)
Call SetTextbox("Due on","BSEG-ZFBDT","",ConvertDate(DT_F22_0303_DUE_ON),False)
'Call SetTextbox("Bus\. Area","BSEG-GSBER","",DT_F22_301_BUS_AREA,False)
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_F22_0303_ASSIGNMENT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F22_0303_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F22_0303_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F22_0303_ACCOUNT,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-22").SapGuiWindow("transaction:=F-22").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F22_0300_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F22_0300_TAX_CODE,False)
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_F22_0300_ASSIGNMENT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F22_0300_TEXT,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_F22_1006_BUSINESS_AREA,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_F22_1006_COST_CENTER,False)
Call SetTextbox("Profit Center","COBL-PRCTR","",DT_F22_1006_PROFIT_CENTER,False)
'Capture the screenshot
Call TakeScreenShot()

'Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
''Capture the screenshot
'Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-22").SapGuiWindow("transaction:=F-22").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next

'veryfy sattus bar content
Call GetStatusBar("item1","DT_DOC_NO_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_F22_0100_CHECK_TEXT_OF_STATUSBAR)
'
'''----------------------Tcode F-22----------------------------
'
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F22_0100_DOCUMENT_DATE_OCC1),False)
'Call SetTextbox("Posting Date","BKPF-BUDAT","",Replace((DT_F22_100_POSTING_DATE),"/","."),False)
Call SetTextbox("Type","BKPF-BLART","",DT_F22_0100_TYPE,False)
Call SetTextbox("Period","BKPF-MONAT","",DT_F22_0100_PERIOD_OCC1,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F22_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F22_0100_CURRENCYRATE_OCC1,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F22_0100_REFERENCE_OCC1,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F22_0100_DOCHEADER_TEXT_OCC1,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F22_0100_PSTKY_OCC1,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F22_0100_ACCOUNT_OCC1,False)
Call SetTextbox("SGL Ind","RF05A-NEWUM","",DT_F22_0100_SGL_IND_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-22").SapGuiWindow("transaction:=F-22").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next

Call SelectCheckbox("BKPF-XMWST","1",DT_F22_0303_CALCULATE_TAX_OCC1,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F22_0303_AMOUNT_OCC1,False)
Call SetTextbox("Due on","BSEG-ZFBDT","",ConvertDate(DT_F22_0303_DUE_ON_OCC1),False)
'Call SetTextbox("Bus\. Area","BSEG-GSBER","",DT_F22_301_BUS_AREA,False)
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_F22_0303_ASSIGNMENT_OCC1,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F22_0303_TEXT_OCC1,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F22_0303_PSTKY_OCC1,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F22_0303_ACCOUNT_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-22").SapGuiWindow("transaction:=F-22").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F22_0300_AMOUNT_OCC1,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F22_0300_TAX_CODE_OCC1,False)
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_F22_0300_ASSIGNMENT_OCC1,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F22_0300_TEXT_OCC1,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_F22_1006_BUSINESS_AREA_OCC1,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_F22_1006_COST_CENTER_OCC1,False)
Call SetTextbox("Profit Center","COBL-PRCTR","",DT_F22_1006_PROFIT_CENTER_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()

'Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
''Capture the screenshot
'Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-22").SapGuiWindow("transaction:=F-22").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next

'veryfy sattus bar content
Call GetStatusBar("item1","DT_DOC_NO_2_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_F22_0100_CHECK_TEXT_OF_STATUSBAR_OCC1)

Wait(1)
Call ClickButton("Back   \(F3\)",False)
Call ClickButtonIfExist("Yes",True)
'Capture the screenshot
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)
'
'''----------------------Tcode F.13----------------------------
'
'Enter the Tcode
Call SetTcode(DT_F22_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Company Code","BUKRX-LOW","",DT_F22_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","GJAHX-LOW","",DT_F22_1000_FISCAL_YEAR,False)
Call SetTextbox("Document Number","DOCNR-LOW","",DT_F22_3010_TABLECELL_SINGLE_VALUE_0,False)
Call SetTextbox("to","DOCNR-HIGH","",DT_F22_3010_TABLECELL_SINGLE_VALUE_1,False)
Call SetTextbox("Customers","KONTD-LOW","",DT_F22_1000_CUSTOMERS,False)
Call SetTextbox("Special G/L Indicator-Customer","SHBKD-LOW","",DT_F22_1000_SPECIAL_GL_INDICATORCUSTOMER,False)

Call SelectCheckbox("XTOL","1",DT_F22_1000_INCLUDE_TOLERANCES,False)
Call SelectCheckbox("X_KUNNR","1",DT_F22_1000_SELECT_CUSTOMERS,False)
Call SelectCheckbox("X_TESTL","1",DT_F22_1000_TEST_RUN,False)
Call SelectCheckbox("X_SHBKN","1",DT_F22_1000_SPECIAL_GL_TRANSACTIONS,False)'edit index
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()   
Wait(5)
'Capture the screenshot
Call TakeScreenShot()

'Validate If doc is generated
Call GetStatusBar("item1","DT_DOC_NO_3_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_F22_0120_CHECK_TEXT_OF_STATUSBAR)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function VerifyifGuiLabelWithIndexExists(Content,Index)
	
 If Not (Environment.Value("blnFatalError") or Content= DS_SKIP) Then
	If blnShowNotification Then BalloonTooltip.Show "TASE Automation","Executing Now : VerifyifGuiLabelWithIndexExists"
	
	strStepName = "Verify if Gui Label exists with index"

    If Content <>"" Then
     set   objLabel = SAPGuiSession(sessionObject).SapGuiWindow(windowobject).SAPGuiLabel("content:="&Content,"guicomponenttype:=30","index:="&Index)
               If objLabel.Exist Then
					Call ReporterFunction(strLibraryFileName,"VerifyifGuiLabelWithIndexExists","2",Content,"Gui Label with value "&Content &" index : "&Index &" exists in the screen")	
					strStatus = "PASS"
					strMsg = "Gui Label with value "&Content&"index : "&Index &" exists in the screen"	
					blnCaptureFlag = True
			              	If blnCaptureFlag  or  blnCreateImageEachStep or blnCreateTrainingDoc Then
		            	  		ImagePath=CaptureScreenshot(strStepName,objLabel,False,False,False)
                              End If
					
					Else
					Call ReporterFunction(strLibraryFileName,"VerifyifGuiLabelWithIndexExists","1"," Status bar Content","Gui Label with value "&Content &"index : "&Index &" doesn't exist  in the screen")	
					strStatus = "FAIL"
					blndefectFlag =True
					strMsg = "Gui Label with value "&Content&" doesn't exist  in the screen"
					blnObjectError=True
				End If
       	Else
    	Call ReporterFunction(strLibraryFileName,"VerifyifGuiLabelWithIndexExists","1","Gui Label","Function Parameter Not Passed Properly. Check the --VerifyifGuiLabelWithIndexExists-- Function Call")
			strStatus = "FAIL" 
			strMsg = "Function Parameter Not Passed Properly. Check the --VerifyifGuiLabelWithIndexExists-- Function Call-"
	End if


If strStatus = "FAIL"  Then
		VerifyifGuiLabelWithIndexExists = strMsg
		blnMainFailFlag = True
		ImagePath=CaptureScreenshot(strStepName,objLabel,False,False,False)
    
	Else
		VerifyifGuiLabelWithIndexExists = True
	End If
	If blnDefault_eSwiftReporting Then  
		Call UpdateResultHtml(strStepName,Content,strMsg,strStatus,"")
	End If

End If
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Call VerifyifGuiLabelExists(DT_F22_0120_CHECK_TEXT_OF_NO_NAME)
Call VerifyifGuiLabelExists(DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC1)
Call VerifyifGuiLabelWithIndexExists(DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC2,"1")
Call VerifyifGuiLabelWithIndexExists(DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC3,"2")
Call VerifyifGuiLabelExists(DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC4)
Call VerifyifGuiLabelExists(DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC5)
Call VerifyifGuiLabelExists(DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC6)

Wait(1)
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)

'
'''----------------------Tcode FB03----------------------------
'
'Enter the Tcode
Call SetTcode(DT_F22_0100_OKCD_OCC1) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Number","RF05L-BELNR","",DT_F22_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_F22_0100_COMPANY_CODE_OCC2,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_F22_0100_FISCAL_YEAR,False)

Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

Call VerifyGridCellContent("",1,"BSCHL","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",2,"BSCHL","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

Call DoubleClickGuiGridCell("",0,1, "Company code", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Account","BSEG-HKONT",0,DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT, False)
Call ClickButton("Back   \(F3\)", False)
Call DoubleClickGuiGridCell("",0,2, "Company code", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Acc","BSEG-HKONT",0,DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT, False)
Call ClickButton("Back   \(F3\)", False)
'Call VerifyGridCellContent("",1,"HKONT","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
'Call VerifyGridCellContent("",2,"HKONT","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)

Call VerifyGridCellContent("",1,"KTONR","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("",2,"KTONR","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)

Call VerifyGridCellContent("",1,"AZBET","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("",2,"AZBET","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)

Call VerifyGridCellContent("",1,"Currency","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_RF05A_UBAZW)
Call VerifyGridCellContent("",2,"Currency","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_RF05A_UBAZW)

Call ClickButton("Display Document Header   \(F5\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call VerifyTextBoxContent("Document type","BKPF-BLART","",DT_F22_1710_CHECK_TEXT_OF_DOCUMENT_TYPE,False)
Call ClickButton("Continue/Confirm   \(Enter\)",True)

Wait(1)
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC5)

'''----------------------Tcode FBL5N----------------------------
''Enter the Tcode
Call SetTcode(DT_F22_0100_OKCD_OCC2) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC6)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_F22_1000_CUSTOMER_ACCOUNT,False)

Call ClickButtonIfExist("Dynamic selections   \(Shift\+F4\)",False)
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_F22_1000_COMPANY_CODE_OCC1,False)

Call ActivateNodeGuiTree("","Customer master;Industry")
Call ActivateNodeGuiTree("","Customer master;Group key")
Call ActivateNodeGuiTree("","Customer master;Country")
Call ActivateNodeGuiTree("","Customer master;Postal Code")
Call ActivateNodeGuiTree("","Customer master;City")
Call ActivateNodeGuiTree("","Customer master;Trading partner")
Call ActivateNodeGuiTree("","Documents;Special G/L ind.")
Call ActivateNodeGuiTree("","Documents;Assignment")
Call ActivateNodeGuiTree("","Company code;Clerk Abbreviation")
Call ActivateNodeGuiTree("","Company code;Reconciliation acct")
'Capture the screenshot
Call TakeScreenShot()

wait(1)
SapGuiSession("transaction:=FBL5N").SapGuiWindow("transaction:=FBL5N").SAPGuiButton("tooltip:=Multiple selection","index:=4").Click

'Capture the screenshot
Call TakeScreenShot()

'set filter criteria
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",1,DT_F22_3010_TABLECELL_SINGLE_VALUE_0_OCC1,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",2,DT_F22_3010_TABLECELL_SINGLE_VALUE_1_OCC1,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",3,DT_F22_3010_TABLECELL_SINGLE_VALUE_2,True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Copy   \(F8\)",True)
'Capture the screenshot
Call TakeScreenShot()

Call SelectRadioButton("X_AISEL","All items",False)
Call SelectCheckbox("X_SHBV","1",DT_F22_1000_SPECIAL_GL_TRANSACTIONS_OCC1,False)

Call ClickButton("Execute   \(F8\)",False)
Wait(2)
Call TakeScreenShot()

''''''''''''Additional code to filter out the correct details as layout got changed to "grid type fro label format'''''''''''
Call ClickButton("Change layout\.\.\.   \(Ctrl\+F8\)",False)
Call SelectRowGuiTableByRow("SAPLSKBHTC_FIELD_LIST",3, True)
Call ClickButton("Show Sel. Fields \(Ctrl\+F3\)",True)
Call SelectRowGuiTableByRow("SAPLSKBHTC_FIELD_LIST",14, True)
Call ClickButton("Show Sel. Fields \(Ctrl\+F3\)",True)
Call SelectRowGuiTableByRow("SAPLSKBHTC_FIELD_LIST",41, True)
Call ClickButton("Show Sel. Fields \(Ctrl\+F3\)",True)
Call ClickButton("Copy   \(Enter\)",True)
Call TakeScreenShot()
'Call ClickButton("Change layout\.\.\.   \(Ctrl\+F8\)",False)
'Call SelectRowGuiGridbyRowNo("Column Set","",46,True)
'Call ClickButton("Show Selected Fields \(F7\)",True)
'Call ClickButton("Transfer   \(Enter\)",True)
'Call TakeScreenShot()

'verify the details
Call VerifyifGuiLabelExists_ByIndex("S_LEDG",0)
Call VerifyifGuiLabelExists_ByIndex("S_LEDG",1)
Call VerifyifGuiLabelExists_ByIndex("S_LEDG",2)
'Call VerifyGridCellContent("",1,"Cleared/open items symbol","",DT_F22_0120_CHECK_TOOLTIP_OF_NO_NAME)
'Call VerifyGridCellContent("",2,"Cleared/open items symbol","",DT_F22_0120_CHECK_TOOLTIP_OF_NO_NAME_OCC1)
'Call VerifyGridCellContent("",3,"Cleared/open items symbol","",DT_F22_0120_CHECK_TOOLTIP_OF_NO_NAME_OCC2)
Call VerifyifGuiLabelExists_ByIndex(DT_F22_0120_CHECK_TEXT_OF_RON,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC7,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC8,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC9,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC10,0)

Call VerifyifGuiLabelExists_ByIndex(DT_F22_0120_CHECK_TEXT_OF_E,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F22_0120_CHECK_TEXT_OF_E_OCC1,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F22_0120_CHECK_TEXT_OF_E_OCC2,0)
'Call VerifyGridCellContent("",1,"Special G/L ind.","",DT_F22_0120_CHECK_TEXT_OF_E)
'Call VerifyGridCellContent("",2,"Special G/L ind.","",DT_F22_0120_CHECK_TEXT_OF_E_OCC1)
'Call VerifyGridCellContent("",3,"Special G/L ind.","",DT_F22_0120_CHECK_TEXT_OF_E_OCC2)
'Call VerifyifGuiLabelExists_ByIndex(DT_F22_0120_CHECK_TEXT_OF_ACCOUNT_11003979,0)

'Call VerifyGridCellContent("",3,"Amount in local currency","",DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC7)
'Call VerifyGridCellContent("",1,"Amount in local currency","",DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC8)
'Call VerifyGridCellContent("",2,"Amount in local currency","",DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC9)
'Call VerifyGridCellContent("",4,"Amount in local currency","",DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC10)

'Call VerifyGridCellContent("",3,"Local Currency","",DT_F22_0120_CHECK_TEXT_OF_RON)

Wait(1)

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()




'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [8,3437477]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//




