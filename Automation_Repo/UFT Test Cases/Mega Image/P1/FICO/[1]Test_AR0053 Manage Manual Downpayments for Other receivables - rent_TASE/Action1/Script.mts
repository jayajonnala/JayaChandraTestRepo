
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_AR0053 Manage Manual Downpayments for Other receivables - rent_TASE
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


gstrTestCaseName = "Test_AR0053 _TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\FICO\TASE_DT_AR0001 Manage AR Documents-Manage Manual Customer Invoicing_Cred.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'''''--------------login----------------'''''
'
''''''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
'
'
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
'
''''----------------------Tcode OB08----------------------------
''Enter the Tcode
Call SetTcode(DT_FD03_0100_OKCD_OCC4) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC10)
'Capture the screenshot
Call TakeScreenShot()

''No_Of_Record = Replace(SAPGuiSession("transaction:=OB08").SAPGuiWindow("transaction:=OB08").SAPGuiEdit("name:=VIM_POSITION_INFO").GetROProperty("value"),"Entry 1","")
'
Call ClickButton("VIM_POSI_PUSH",False)

Call SetTextboxNolabel("SVALD-VALUE","",DT_FD03_0300_EXCH_RATE_TYPE,True)
Call SetTextbox("From currency","SVALD-VALUE","",DT_FD03_0300_FROM_CURRENCY,True)
Call SetTextbox("To-currency","SVALD-VALUE","",DT_FD03_0300_TOCURRENCY,True)
Call SetTextbox("Valid from","SVALD-VALUE","",DT_FD03_0300_VALID_FROM,True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)
'Capture the screenshot
Call TakeScreenShot()

'Posi_Row = Replace(Replace(Replace(SAPGuiSession("transaction:=OB08").SAPGuiWindow("transaction:=OB08").SAPGuiEdit("name:=VIM_POSITION_INFO").GetROProperty("value"),No_Of_Record,""),"Entry ",""),".","")

Call GetTextBoxvalue("VIM_POSITION_INFO",0,"DT_ROW_OUTPUT_2",False)
''Call GetTableCellData("SAPL0SAPTCTRL_V_TCURR","Dir.quot.",Posi_Row,"ValidFrom",DT_FD03_0300_VALID_FROM,"DT_FD03_0020_CHECK_TEXT_OF_TABLECELL_DIRQUOT_0_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call GetTableCellData("SAPL0SAPTCTRL_V_TCURR","Dir.quot.",DT_ROW_EXRT,"From","EUR","DT_FD03_0020_CHECK_TEXT_OF_TABLECELL_DIRQUOT_0_OUTPUT",False)


Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call ClickButton("Back   \(F3\)",False)
'
'''----------------------Tcode FD03----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Customer","RF02D-KUNNR","",DT_FD03_7106_CUSTOMER,True)
Call SetTextbox("Company Code","RF02D-BUKRS","",DT_FD03_7106_COMPANY_CODE,True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)

wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyTextBoxContent("Customer","RF02D-KUNNR","",DT_FD03_7002_CHECK_TEXT_OF_CUSTOMER,False)
Call ClickButton("Company Code Data   \(Ctrl\+F2\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call VerifyTextBoxContent("Recon\. account","KNB1-AKONT","",DT_FD03_7211_CHECK_TEXT_OF_RECON_ACCOUNT,False)

Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Cancel   \(F12\)",True)

'Capture the screenshot
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)
'
'''''----------------------Tcode FB70----------------------------
'Enter the Tcode
Call SetTcode(DT_FD03_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_FD03_1000_COMPANY_CODE)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)

Call SetTextbox("Customer","INVFO-ACCNT","",DT_FD03_0510_CUSTOMER,False)
Call SetTextbox("Invoice date","INVFO-BLDAT","",Replace((DT_FD03_0510_INVOICE_DATE),"/","."),False)
Call SetTextbox("Posting Date","INVFO-BUDAT","",Replace((DT_FD03_0510_POSTING_DATE),"/","."),False)
Call SetTextbox("Reference","INVFO-XBLNR","",DT_FD03_0510_REFERENCE,False)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FD03_0510_AMOUNT,False)
Call SetTextbox("Amount","INVFO-WAERS","",DT_FD03_0510_AMOUNT_OCC1,False)
'Select Calculate Tax field as true
Call SelectCheckbox("INVFO-XMWST",0,DT_FD03_0510_CALCULATE_TAX,False)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
'
If SAPGuiSession("transaction:=FB70").SapGuiWindow("transaction:=FB70").SAPGuiStatusBar("messagetype:=W").Exist = True Then
	Call PressEnter() 
End If 
Call TakeScreenShot()
Call SetComboByKey("INVFO-MWSKZ",DT_FD03_0510_INVFOMWSKZ)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 
For Iterator = 1 To 10 Step 1
	If SAPGuiSession("transaction:=FB70").SapGuiWindow("transaction:=FB70").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next

''Capture the screenshot
Call TakeScreenShot()
Call SetTextbox("HeaderText","INVFO-BKTXT","",DT_FD03_0550_HEADERTEXT,False)
Call PressEnter() 

'set table data
Call SetTableDataNoRef("SAPLFSKBTABLE","G/L acct",1,DT_FD03_0100_TABLECELL_GL_ACCT_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Amount in doc.curr.",1,DT_FD03_0100_TABLECELL_AMOUNT_IN_DOCCURR_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Text",1,DT_FD03_0100_TABLECELL_TEXT_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Business area",1,"",False)'DT_FD03_0100_TABLECELL_BUSINESS_AREA_0
Call SetTableDataNoRef("SAPLFSKBTABLE","Cost center",1,"",False)'DT_FD03_0100_TABLECELL_COST_CENTER_0
Call SetTableDataNoRef("SAPLFSKBTABLE","Profit center",1,DT_PROFIT_CENTER,False)
Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

'Call SelectTab("TS",DT_ZFIAR_RS_RFKORD50PDF_1200_DETAILS,False)
''Capture the screenshot
'Call TakeScreenShot()
'Call SelectTab("TS",DT_ZFIAR_RS_RFKORD50PDF_1200_TAX,False)
''Capture the screenshot
'Call TakeScreenShot()
'Call SelectTab("TS",DT_ZFIAR_RS_RFKORD50PDF_1200_AMOUNT_SPLIT,False)
''Capture the screenshot
'Call TakeScreenShot()
'Call SelectTab("TS",DT_ZFIAR_RS_RFKORD50PDF_1200_NOTES,False)
''Capture the screenshot
'Call TakeScreenShot()
'Call SetTextArea(DT_ZFIAR_RS_RFKORD50PDF_0540_TEXTEDIT_SHELL)

Call SelectMenuBar("Document;Simulate")
Wait(2)
Call PressEnter() 
Wait(2)

'Capture the screenshot
Call TakeScreenShot()

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
Call VerifyifGuiLabelExists(DT_FD03_0120_CHECK_TEXT_OF_NO_NAME) 
Call VerifyifGuiLabelExists(DT_FD03_0120_CHECK_TEXT_OF_NO_NAME_OCC1) 
Call VerifyifGuiLabelExists(DT_FD03_0120_CHECK_TEXT_OF_NO_NAME_OCC2) 

Call VerifyifGuiLabelExists(DT_FD03_0120_CHECK_TEXT_OF_NO_NAME_OCC3) 
Call VerifyifGuiLabelExists(DT_FD03_0120_CHECK_TEXT_OF_NO_NAME_OCC4) 
Call VerifyifGuiLabelExists(DT_FD03_0120_CHECK_TEXT_OF_NO_NAME_OCC5) 

'Click on Post Buton
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
'Validate If document is posted and get the status bar nummber
Call GetStatusBar("item1","DT_DOC_NO_1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_CHECK_TEXT_OF_STATUSBAR)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Cancel   \(F12\)",True)

'Capture the screenshot
Call TakeScreenShot()
'''Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)

''----------------------Tcode FB03----------------------------

'Enter the Tcode
Call SetTcode(DT_FD03_0100_OKCD_OCC1) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Number","RF05L-BELNR","",DT_FD03_0100_DOCUMENT_NUMBER1,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FD03_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FD03_0100_FISCAL_YEAR,False)

Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

Call GetGridContentByTitle("","","Assignment",1,"DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call DoubleClickGuiGridCell("",0, 1, "Company code", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Acc","BSEG-HKONT",0,DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT, False)
Call ClickButton("Back   \(F3\)", False)
'Call VerifyGridCellContent("",1,"HKONT","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
Call DoubleClickGuiGridCell("",0,2, "Company code", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Account","BSEG-HKONT",0,DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT, False)
Call ClickButton("Back   \(F3\)", False)
'Call VerifyGridCellContent("",2,"HKONT","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)
Call DoubleClickGuiGridCell("",0,3, "Company code", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Account","BSEG-HKONT",0,DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_HKONT, False)
Call ClickButton("Back   \(F3\)", False)
'Call VerifyGridCellContent("",3,"HKONT","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_HKONT)

Call VerifyGridCellContent("",1,"BSCHL","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",2,"BSCHL","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("",3,"BSCHL","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)

Call VerifyGridCellContent("",1,"Account","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("",2,"Account","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("",3,"Account","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)

Call VerifyGridCellContent("",1,"AZBET","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("",2,"AZBET","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("",3,"AZBET","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AZBET)

Call VerifyGridCellContent("",1,"RF05A_UBAZW","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_RF05A_UBAZW)
Call VerifyGridCellContent("",2,"RF05A_UBAZW","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_RF05A_UBAZW)
Call VerifyGridCellContent("",3,"RF05A_UBAZW","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_RF05A_UBAZW)

Wait(1)
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC5)
'
''''----------------------Tcode FB01----------------------------
''Enter the Tcode
Call SetTcode(DT_FD03_0100_OKCD_OCC2) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC6)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Date","BKPF-BLDAT","",Replace((DT_FD03_0100_DOCUMENT_DATE),"/","."),False)
Call SetTextbox("Type","BKPF-BLART","",DT_FD03_0100_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FD03_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",Replace((DT_FD03_0100_POSTING_DATE),"/","."),False)
Call SetTextbox("Period","BKPF-MONAT","",DT_FD03_0100_PERIOD,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FD03_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_FD03_0100_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_DOC_HEADER_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FD03_0100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FD03_0100_ACCOUNT,False)
'Call SetTextbox("SGL Ind","RF05A-NEWUM","",DT_FB01_0100_SGL_IND,False)
'Capture the screenshot
Call TakeScreenShot()
While SAPGuiSession(sessionObject).SapGuiWindow(windowobject).SAPGuiEdit("guicomponenttype:=31","attachedtext:=Reference","name:=BKPF-XBLNR","Index:=0").Exist(5)=True 
Call PressEnter()  
Wend
'''Call PressEnter() 
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FD03_0300_AMOUNT,False)
'Call SetTextbox("Value date","BSEG-VALUT","",Replace((DT_FB01_0300_VALUE_DATE),"/","."),False)
'Call SetTextbox("Business Area","COBL-GSBER","",DT_FB01_1007_BUSINESS_AREA,False)
'Call SetTextbox("Cost Center","COBL-KOSTL","",DT_FB01_1007_COST_CENTER,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_FD03_0300_TAX_CODE,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_FD03_0300_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FD03_0300_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FD03_0300_ACCOUNT,False)
Call SetTextbox("SGL Ind","RF05A-NEWUM","",DT_FD03_0300_SGL_IND,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FD03_0304_AMOUNT,False)
'Call SetTextbox("Bus\. Area","BSEG-GSBER","",DT_FB01_0301_BUS_AREA,False)
'Call SetTextbox("Bline Date","BSEG-ZFBDT","",Replace((DT_FB01_0301_BLINE_DATE),"/","."),False)
'Capture the screenshot
Call TakeScreenShot()

'Call SelectMenuBar("Document;Simulate")
'Wait(2)
''Capture the screenshot
'Call TakeScreenShot()
'Call SelectMenuBar("Goto;First item")
'Call VerifyTextBoxContent("G/L Account","BSEG-HKONT","",DT_FB01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR,False)
'Call SelectMenuBar("Goto;Next Item")
'Call VerifyTextBoxContent("Customer","KNA1-KUNNR","",DT_FB01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR,False)
'Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
'Call SelectMenuBar("Document;Simulate General Ledger")
'Call VerifyGridCellContent("",1,"Posting Key","",DT_FB01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
'Call VerifyGridCellContent("",2,"Posting Key","",DT_FB01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
'Call ClickButton("Back   \(F3\)",False)
''Capture the screenshot
'Call TakeScreenShot()

'Call ClickButton("Post   \(Ctrl\+S\)",False)
'Call ClickButtonIfExist("Save",True)
'wait(1)
'Call TakeScreenShot()
''Validate If invoice is generated
'Call GetStatusBar("item1","DT_DOC_NO_2_OUTPUT")
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'VerifyStatusBar(DT_FD03_0100_CHECK_TEXT_OF_STATUSBAR)
'
'Wait(1)
'Call ClickButton("Back   \(F3\)",False)
''Capture the screenshot
'Call TakeScreenShot()
'Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC7)
'
''''----------------------Tcode FB03----------------------------
''
''Enter the Tcode
'Call SetTcode(DT_FD03_0100_OKCD_OCC3) 
'Call PressEnter()     ' 
'Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC8)
''Capture the screenshot
'Call TakeScreenShot()
'
'Call SetTextbox("Document Number","RF05L-BELNR","",DT_FD03_0100_DOCUMENT_NUMBER_OCC1,False)
'Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FD03_0100_COMPANY_CODE_OCC2,False)
'Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FD03_0100_FISCAL_YEAR_OCC1,False)
'
'Call PressEnter()
''Capture the screenshot
'Call TakeScreenShot()
'
'Call GetGridContentByTitle("","","Assignment",1,"DT_GET_REF_OUTPUT")
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'Call DoubleClickGuiGridCell("",0, 1, "Company code", False)
'Call TakeScreenShot
'Call VerifyTextBoxContent("G/L Account","BSEG-HKONT",0,DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT_OCC1, False)
'Call ClickButton("Back   \(F3\)", False)
''Call VerifyGridCellContent("",1,"HKONT","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT_OCC1)
'Call DoubleClickGuiGridCell("",0,2, "Company code", False)
'Call TakeScreenShot
'Call VerifyTextBoxContent("G/L acct","BSEG-HKONT",0,DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT_OCC1, False)
'Call ClickButton("Back   \(F3\)", False)
''Call VerifyGridCellContent("",2,"HKONT","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT_OCC1)
'
'Call VerifyGridCellContent("",1,"BSCHL","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OCC1)
'Call VerifyGridCellContent("",2,"BSCHL","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL_OCC1)
'
'Call VerifyGridCellContent("",1,"Account","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR_OCC1)
'
'Call VerifyGridCellContent("",2,"LOKKT","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT)
'
'Call VerifyGridCellContent("",1,"ZUONR","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_OCC1)
'Call VerifyGridCellContent("",2,"ZUONR","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR)
'
'Call VerifyGridCellContent("",1,"AZBET","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET_OCC1)
'Call VerifyGridCellContent("",2,"AZBET","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET_OCC1)
'
'Call VerifyGridCellContent("",1,"RF05A_UBAZW","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_RF05A_UBAZW_OCC1)
'Call VerifyGridCellContent("",2,"RF05A_UBAZW","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_RF05A_UBAZW_OCC1)
'
'Call ClickButton("Change Display Currency   \(F8\)",False)
'SAPGuiSession("Session").SAPGuiWindow("Display Currency for Document").InsightObject("InsightObject").Click
'Call ClickButton("Continue/Confirm   \(Enter\)",True)
'Call VerifyGridCellContent("",1,"Amount","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMBTR)
'Call VerifyGridCellContent("",2,"Amount","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMBTR)
''Call VerifyGridCellContent("",1,"DMBTR","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMBTR)
''Call VerifyGridCellContent("",2,"DMBTR","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMBTR)
'
'
'Wait(1)
'Call ClickButton("Back   \(F3\)",False)
'Call ClickButton("Back   \(F3\)",False)
''Capture the screenshot
'Call TakeScreenShot()
'Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC11)
''
'''----------------------Tcode FBL5N----------------------------
'Enter the Tcode
Call SetTcode(DT_FD03_0100_OKCD_OCC5) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC12)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_FD03_1000_CUSTOMER_ACCOUNT,False)

Call ClickButtonIfExist("Dynamic selections   \(Shift\+F4\)",False)
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_FD03_1000_COMPANY_CODE_OCC1,False)

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
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",1,DT_FD03_3010_TABLECELL_SINGLE_VALUE_0,True)

'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Copy   \(F8\)",True)

Call SelectRadioButton("X_AISEL","All items",False)
Call SetTextbox("Layout","PA_VARI","","1SAP-SBWAP",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Wait(2)
Call TakeScreenShot()
Call VerifyStatusBarMessageType("S")

Call VerifyifGuiLabelWithIndexExists(DT_FD03_0120_CHECK_TEXT_OF_HR3000000288,"0")
Call VerifyifGuiLabelExists(DT_FD03_0120_CHECK_TEXT_OF_NO_NAME_OCC6)
Call VerifyifGuiLabelExists(DT_FD03_0120_CHECK_TEXT_OF_DR)
Call VerifyifGuiLabelWithIndexExists(DT_FD03_0120_CHECK_TOOLTIP_OF_NO_NAME,"1")
Call VerifyifGuiLabelWithIndexExists(DT_FD03_0120_CHECK_TEXT_OF_RON,"1")
Call VerifyifGuiLabelWithIndexExists(DT_FD03_0120_CHECK_TEXT_OF_NO_NAME_OCC7,"1")

Wait(2)
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




