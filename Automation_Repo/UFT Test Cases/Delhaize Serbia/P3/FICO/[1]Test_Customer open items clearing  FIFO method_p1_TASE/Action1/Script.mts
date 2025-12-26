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
'.................Test Script Name : Test_Customer open items clearing FIFO method_p1_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 1st May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Customer open items clearing FIFO method_p1_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Customer open items clearing  FIFO method_p1_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

' StartExecution(excelPath, strTestCaseName, iterationIndex, strResultFolderPath)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath )  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode FBL5N----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

Call SetTextbox("Open at key date","PA_STIDA","",Replace((DT_FBL5N_1000_OPEN_AT_KEY_DATE),"/","."),False)
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_FBL5N_1000_COMPANY_CODE,False)
Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_FBL5N_1000_CUSTOMER_ACCOUNT,False)
Call TakeScreenShot()

Call FocusTextBox("Open at key date","PA_STIDA",False)
Call ClickButton("Execute   \(F8\)",False)
Wait(2)
Call TakeScreenShot()

CAll ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(10)

Call TakeScreenShot()
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)

Call GetTextStatusBar("DT_FBL5N_1000_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

VerifyStatusBar(DT_FBL5N_1000_CHECK_TEXT_OF_STATUSBAR_OCC1)
Call TakeScreenShot()

Call ClickButton("Exit   \(Shift\+F3\)",False)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()


'''----------------------Tcode FB70----------------------------

'Enter the Tcode
Call SetTcode(DT_FBL5N_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_FBL5N_0100_OKCD)
'Capture the screenshot
Call TakeScreenShot()

'Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FBL5N_1000_COMPANY_CODE_OCC1,True)
Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_FBL5N_1000_COMPANY_CODE_OCC1)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Continue   \(Enter\)",True)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Customer","INVFO-ACCNT","",DT_FBL5N_0510_CUSTOMER,False)
Call SetTextbox("Invoice date","INVFO-BLDAT","",Replace((DT_FBL5N_0510_INVOICE_DATE),"/","."),False)
Call SetTextbox("Reference","INVFO-XBLNR","",DT_FBL5N_0510_REFERENCE,False)
Call SetTextbox("Posting Date","INVFO-BUDAT","",Replace((DT_FBL5N_0510_POSTING_DATE),"/","."),False)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FBL5N_0510_AMOUNT,False)
'Select Calculate Tax field as true
Call SelectCheckbox("INVFO-XMWST",0,DT_FBL5N_0510_CALCULATE_TAX,False)

'Capture the screenshot
Call TakeScreenShot()

'set table data
Call SetTableDataNoRef("SAPLFSKBTABLE","G/L acct",1,DT_FBL5N_0100_TABLECELL_GL_ACCT_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Amount in doc.curr.",1,DT_FBL5N_0100_TABLECELL_AMOUNT_IN_DOCCURR_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Tax code",1,DT_FBL5N_0100_TABLECELL_TAX_CODE_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Business area",1,DT_FBL5N_0100_TABLECELL_BUSINESS_AREA_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Cost center",1,DT_FBL5N_0100_TABLECELL_COST_CENTER_0,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TS",DT_FBL5N_1200_DETAILS,False)
'Capture the screenshot
Call TakeScreenShot()
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC2)

Call ClickButton("Enter",False)
'Capture the screenshot
Call TakeScreenShot()
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)

Call SetTextbox("HeaderText","INVFO-BKTXT","",DT_FBL5N_0550_HEADERTEXT,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Simulate Document Posting   \(F9\)",False)
Wait(1)
Call PressEnter()
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
'validate components
Call VerifyifGuiLabelExists(DT_FBL5N_0750_CHECK_TEXT_OF_COMPANY_CODE)
Call VerifyifGuiLabelExists(DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyifGuiLabelWithIndexExists(DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL,0)
Call VerifyifGuiLabelWithIndexExists(DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL,1)
Call VerifyifGuiLabelExists(DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
Call VerifyifGuiLabelExists(DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_HKONT)
Call VerifyifGuiLabelExists(DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_HKONT1)

'Capture the screenshot
Call TakeScreenShot()
'Click on Post Buton
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
'Validate If document is posted and get the status bar nummber
Call GetStatusBar("item1","DT_DOCUMENT_NUMBER_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_FB70_CHECK_TEXT_OF_STATUSBAR_OCC1)
'Capture the screenshot
Call TakeScreenShot()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FBL5N_1000_COMPANY_CODE_OCC2,True)
Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_FBL5N_1000_COMPANY_CODE_OCC2)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Continue   \(Enter\)",True)
'Capture the screenshot
Call TakeScreenShot()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

