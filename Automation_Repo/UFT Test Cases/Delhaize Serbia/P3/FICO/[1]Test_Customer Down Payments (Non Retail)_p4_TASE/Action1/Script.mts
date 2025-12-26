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

 '.......................Mandatory Initial Call only in First Component in a Test Scenario
'reload DS to update dates and calculations
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Customer Down Payments (Non Retail)_p4_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 12th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Customer Down Payments (Non Retail)_p4_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Customer Down Payments (Non Retail)_p4_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath) 

Call EndDateof445PeriodByDate(DT_TODAY,"DT_ENDING_DATE_PERIOD")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'Increment the parameter
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_REFERENCE",(Cint(DT_INCREMENT_REFERENCE)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''
'''----------------------Tcode ZFIAR_RS_RFKORD50PDF----------------------------
'''Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Created By","ENAME-LOW","","",True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",True)
'Capture the screenshot
Call TakeScreenShot()
Call SelectRowGuiGridbyRowNo("Variant Catalog for Program ZFIAR_RS_RFKORD50_PDF","",DT_ZFIAR_RS_RFKORD50PDF_0600_GRIDCELL_5_SHORT_DESCRIPTION,False)
Call ClickButton("Choose   \(F2\)",True)

Call SetTextbox("Document Number","RBELNR-LOW","",DT_ZFIAR_RS_RFKORD50PDF_1000_DOCUMENT_NUMBER,False)
Call FocusTextBox("Document Number","RBELNR-LOW",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Wait(5)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Exit   \(Shift\+F3\)",False)
Call ClickButtonIfExist("Exit   \(Shift\+F3\)",False)
''
'''----------------------Tcode FB70----------------------------
'Enter the Tcode
Call SetTcode(DT_ZFIAR_RS_RFKORD50PDF_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_ZFIAR_RS_RFKORD50PDF_0100_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_Company_Code)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)

Call ClickButtonIfExist("Switch Company Code   \(F7\)",False)
Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_Company_Code)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)

Call SetTextbox("Customer","INVFO-ACCNT","",DT_ZFIAR_RS_RFKORD50PDF_0510_CUSTOMER,False)
Call PressEnter() 
Call SetTextbox("Invoice date","INVFO-BLDAT","",Replace((DT_ZFIAR_RS_RFKORD50PDF_0510_INVOICE_DATE),"/","."),False)
Call SetTextbox("Posting Date","INVFO-BUDAT","",Replace((DT_ZFIAR_RS_RFKORD50PDF_0510_POSTING_DATE),"/","."),False)
Call PressEnter() 
Call PressEnter()
Call PressEnter()
Call SetTextbox("Reference","INVFO-XBLNR","",DT_ZFIAR_RS_RFKORD50PDF_0510_REFERENCE,False)
Call PressEnter() 
Call PressEnter()
Call SetTextbox("HeaderText","INVFO-BKTXT","",DT_ZFIAR_RS_RFKORD50PDF_0550_HEADERTEXT,False)
Call PressEnter() 
Call SelectTab("TS","Basic data",False)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_ZFIAR_RS_RFKORD50PDF_0510_AMOUNT,False)
'Select Calculate Tax field as true
Call SelectCheckbox("INVFO-XMWST",0,DT_ZFIAR_RS_RFKORD50PDF_0510_CALCULATE_TAX,False)
Call PressEnter() 
Call PressEnter() 

'set table data
Call SetTableDataNoRef("SAPLFSKBTABLE","G/L acct",1,DT_ZFIAR_RS_RFKORD50PDF_0100_TABLECELL_GL_ACCT_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Amount in doc.curr.",1,DT_ZFIAR_RS_RFKORD50PDF_0100_TABLECELL_AMOUNT_IN_DOCCURR_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Tax code",1,DT_ZFIAR_RS_RFKORD50PDF_0100_TABLECELL_TAX_CODE_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Business area",1,DT_ZFIAR_RS_RFKORD50PDF_0100_TABLECELL_BUSINESS_AREA_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Cost center",1,DT_ZFIAR_RS_RFKORD50PDF_0100_TABLECELL_COST_CENTER_0,False)
Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TS",DT_ZFIAR_RS_RFKORD50PDF_1200_DETAILS,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TS",DT_ZFIAR_RS_RFKORD50PDF_1200_TAX,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TS",DT_ZFIAR_RS_RFKORD50PDF_1200_AMOUNT_SPLIT,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TS",DT_ZFIAR_RS_RFKORD50PDF_1200_NOTES,False)
'Capture the screenshot
Call TakeScreenShot()
Call SetTextArea(DT_ZFIAR_RS_RFKORD50PDF_0540_TEXTEDIT_SHELL)

CAll SelectMenuBar("Document;Simulate")
Wait(2)
Call ClickButtonIfExist("Simulate Document Posting   \(F9\)",False)
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
Call VerifyifGuiLabelExistsByRelativeid(DT_ZFIAR_RS_RFKORD50PDF_0120_CHECK_TEXT_OF_NO_NAME,"wnd\[0\]/usr/lbl\[5,13\]")
'''Call VerifyifGuiLabelExists(DT_ZFIAR_RS_RFKORD50PDF_0120_CHECK_TEXT_OF_NO_NAME) 
Call VerifyifGuiLabelWithIndexExists(DT_ZFIAR_RS_RFKORD50PDF_0120_CHECK_TEXT_OF_NO_NAME_OCC1,"0") 
Call VerifyifGuiLabelWithIndexExists(DT_ZFIAR_RS_RFKORD50PDF_0120_CHECK_TEXT_OF_NO_NAME_OCC2,"1") 
Call VerifyifGuiLabelExists(DT_ZFIAR_RS_RFKORD50PDF_0120_CHECK_TEXT_OF_NO_NAME_OCC3) 
Call VerifyifGuiLabelExists(DT_ZFIAR_RS_RFKORD50PDF_0120_CHECK_TEXT_OF_NO_NAME_OCC4) 
Call VerifyifGuiLabelExists(DT_ZFIAR_RS_RFKORD50PDF_0120_CHECK_TEXT_OF_NO_NAME_OCC5) 
Call VerifyifGuiLabelWithIndexExists(DT_ZFIAR_RS_RFKORD50PDF_0120_CHECK_TEXT_OF_MS,"0") 
Call VerifyifGuiLabelWithIndexExists(DT_ZFIAR_RS_RFKORD50PDF_0120_CHECK_TEXT_OF_MS_OCC1,"1") 
Call VerifyifGuiLabelWithIndexExists(DT_ZFIAR_RS_RFKORD50PDF_0120_CHECK_TEXT_OF_MS_OCC2,"2") 
Call VerifyifGuiLabelExists(DT_ZFIAR_RS_RFKORD50PDF_0120_CHECK_TEXT_OF_NO_NAME_OCC6) 
Call VerifyifGuiLabelExists(DT_ZFIAR_RS_RFKORD50PDF_0120_CHECK_TEXT_OF_NO_NAME_OCC7) 
Call VerifyifGuiLabelExists(DT_ZFIAR_RS_RFKORD50PDF_0120_CHECK_TEXT_OF_NO_NAME_OCC8) 
  

'Click on Post Buton
Call ClickButton("Post   \(Ctrl\+S\)",False)
wait 2
'Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
'''Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
'''Validate If document is posted and get the status bar nummber
Call GetStatusBar("item1","DT_ZFIAR_RS_RFKORD50PDF_1200_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetStatusBar("item2","DT_ZFIAR_RS_RFKORD50PDF_1200_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_ZFIAR_RS_RFKORD50PDF_1200_CHECK_TEXT_OF_STATUSBAR)
'''Capture the screenshot
Call TakeScreenShot()


'''Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

