
'''''''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'''''''.................Test Script Name :Test_09.05.01.10.14 Checks encashment through FCHR_V2
'''''''.................Author : TCS 
'''''''................ Creation Date :
'''''''.................Modified By :
'''''''.................Modified Date/Details :
'''''''
'''''''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
''''''

If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	GetRowNo= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_09.05.01.10.14 Checks encashment through FCHR_V2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'''gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


''''''''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'''''DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''--------TransactionCode-f-58 ----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Company Code","OPAYF-BUKRS","",DT_F58_0130_COMPANY_CODE,False)
Call SetTextbox("Payment method","OPAYF-RZAWE","",DT_F58_0130_PAYMENT_METHOD,False)
Call SetTextbox("House bank","OPAYF-HBKID","",DT_F58_0130_HOUSE_BANK,False)
Call SetTextbox("Account ID","OPAYF-HKTID","",DT_F58_0130_OPAYFHKTID,False)
Call SetTextbox("Check lot number","OPAYF-PSTAP","",DT_F58_0130_CHECK_LOT_NUMBER,False)
Call SetTextbox("Printer for forms","OPAYF-PPRIZ","",DT_F58_0130_PRINTER_TO_FORMS,False)
Call TakeScreenShot
Call PRessEnter()
Call ClickButtonIfExist("Enter Payments   \(Shift\+F1\)",False)
Call TakeScreenShot
Call SelectCheckbox("OPAYF-XAKOB", 0,DT_F58_0129_PAYMENT_ON_ACCT , False)
Call TakeScreenShot

Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F58_0129_DOCUMENT_DATE),False)
Call SetTextbox("Type","BKPF-BLART","",DT_F58_0129_TYPE,False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_F58_0129_POSTING_DATE),False)
Call SetTextbox("Period","BKPF-MONAT","",Cstr(Month(DT_F58_0129_PERIOD)),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F58_0129_CURRENCYRATE,False)
Call SetTextbox("Value date","BSEG-VALUT","",ConvertDate(DT_F58_0129_VALUE_DATE),False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F58_0129_TEXT,False)
Call SetTextbox("Pmnt on acct","OPAYF-SGTXT","",DT_F58_0129_PMNT_ON_ACCT,False)
'''Call SetTextbox("Vendor","OPAYF-LIFNR","",DT_F58_0129_VENDOR,False)
Call SetTextbox("Supplier","OPAYF-LIFNR","",DT_F58_0129_VENDOR,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F58_0129_AMOUNT,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SelectMEnuBar("Document;Simulate")
Call TakeScreenShot


Call FocusTextBoxByIndex("PK  BusA Acct                               EUR   Amount        Tax amnt","RF05A-AZEI1",0,False)
Call ClickButtonIfExist("Choose   (F2)",False)
'Call Sendkey("{F2}")
Wait(3)
Call SetTextbox("Profit Ctrs","COBL-PRCTR","","GC000001",False)
Call Clickbutton("Display Additional Data for Document Item   \(F7\)",False)
Call TakeScreenShot
Call SetTextbox("House Bank","BSEG-HBKID","",DT_F58_0330_HOUSE_BANK,False)
Call SetTextbox("/","BSEG-HKTID","",DT_F58_0330_BSEGHKTID,False)
Call PressEnter()
Call TakeScreenShot

Call SelectMenuBar("Document;Simulate")
Call TakeScreenShot

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot

Call GetTextboxValue("MESSTXT1", 0, "DT_F58_0010_CHECK_TEXT_OF_MESSTXT1_OUTPUT", True)
Call WriteRunTimeDataToExcelGlobalSheet("DT_F58_0010_CHECK_TEXT_OF_MESSTXT1_OUTPUT",DT_F58_0010_CHECK_TEXT_OF_MESSTXT1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call ClickButton("Continue   \(Enter\)",True)

Call SetFocusGuiLabel("Outgoing checks", 11, 88, False)
Call ClickButtonIfExist("Choose   (F2)",False)
'Call Sendkey("{F2}")
Wait(3)
Call TakeScreenShot
Call SetFocusGuiLabel(DT_ENVIRONMENT, 25, 56, False)
Call DoubleClick()
'Call Sendkey("{F2}")
Wait(3)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)
Wait(3)
Call ClickButton("Back   \(F3\)",False)
Wait(3)
Call ClickButton("Back   \(F3\)",False)
Wait(3)
Call ClickButton("Back   \(F3\)",False)
Wait(3)
Call ClickButton("Delete   \(Shift\+F2\)",False)
Call ClickButton("Exit   \(Shift\+F3\)",False)


'''''--------TransactionCode-FCHN ----------''''
Call SetTcode(DT_F58_0120_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Paying Company Code","SEL_ZBUK-LOW","",DT_F58_1000_PAYING_COMPANY_CODE,False)
Call SetTextbox("House bank","SEL_HBKI-LOW","",DT_F58_1000_HOUSE_BANK,False)
Call SetTextbox("Account ID","SEL_HKTI-LOW","",DT_F58_1000_ACCOUNT_ID,False)
Call TakeScreenShot

Call SelectTab("TABSTRIP_CHK","Further Selections", False)
Call TakeScreenShot
Call SetTextbox("Payment document no\.","SEL_VBLN-LOW","",DT_F58_0002_PAYMENT_DOCUMENT_NO,False)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot

Call ClickButton("Execute   \(F8\)",False)

Call GetLabelContentByRefLabel("Outgoing checks", -7,-64, "DT_F58_0120_CHECK_TEXT_OF_NO_NAME_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_F58_0120_CHECK_TEXT_OF_NO_NAME_OUTPUT",DT_F58_0120_CHECK_TEXT_OF_NO_NAME)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call LogOff'
Call FinalStatus()













Public Function FocusTextBoxByIndex(attachedText,textboxName,textboxIndex,blnIsItPopup)

If Not (Environment.Value("blnFatalError")) Then
	   If blnShowNotification Then BalloonTooltip.Show "TASE Automation","Executing Now : FocusTextBoxByIndex"
Dim objTextboxGuiComp32,objTextboxGuiComp31,objTextbox,objWindow
strStepName=" Set Focus on the TextField " &textboxName
If attachedText<>"" and textboxName<>"" then
Set objWindow= SetSAPwindowObj(blnIsItPopup)
	Set objTextboxGuiComp32 = SAPGuisession(sessionObject).sapguiwindow(objWindow).sapguiedit("guicomponenttype:=32","attachedtext:="&attachedText,"name:="&textboxName,"index:="&textboxIndex)
	Set objTextboxGuiComp31 = SAPGuisession(sessionObject).sapguiwindow(objWindow).sapguiedit("guicomponenttype:=31","attachedtext:="&attachedText,"name:="&textboxName,"index:="&textboxIndex)
	Set objTextbox=SetObj(objTextboxGuiComp32,objTextboxGuiComp31)
	
	If objTextbox.Exist Then
				
	      	  If attachedText<>""  then
				strStepName=" Set Focus on the TextField " &attachedText    
				else      
				strStepName=" Set Focus on the TextField " &textboxName   			  
		      End If
    				objTextbox.SetFocus
    				strStatus = "DONE"
    				If blnCaptureFlag  or  blnCreateImageEachStep or blnCreateTrainingDoc Then
				  		ImagePath=CaptureScreenshot(strStepName,objTextbox,False,False,False)
	          		End if
           		  strMsg=" Textbox Focused successfully"
           		  Call ReporterFunction(strLibraryFileName,"FocusTextBoxByIndex","2","Text Box field","Textbox Focused successfully")
           		  
           		  
	Else
			Call ReporterFunction(strLibraryFileName,"FocusTextBoxByIndex","1","Text Box field","Object Missing. Textbox Not Found")
			strStatus="FAIL"
			strMsg = "Object Missing. Textbox Not Found"
			blnObjectError=True
	End if	
else	
Call ReporterFunction(strLibraryFileName,"FocusTextBoxByIndex","1","Text Box field","Function Parameter Not Passed Properly. Check the --FocusTextBoxByIndex-- Function Call")
			strStatus="FAIL"
			strMsg = "Function Parameter Not Passed Properly. Check the --FocusTextBoxByIndex-- Function Call"
		End If
       If  blnObjectError  Then
		    Environment.Value("blnFatalError")=True
	   End If

	   If strStatus = "FAIL"  Then
		    FocusTextBoxByIndex= strMsg
		    blnMainFailFlag = True
		    ImagePath=CaptureScreenshot(strStepName,objTextbox,False,False,False)
	   Else
		    FocusTextBoxByIndex = strMsg
	   End If
	
	   If blnDefault_eSwiftReporting Then  
		   Call UpdateResultHtml(strStepName,"",strMsg,strStatus,"")
	   End If
Set objTextboxGuiComp32=Nothing
Set objTextbox=Nothing
Set objTextboxGuiComp31=Nothing
Set objWindow=Nothing
End If
End Function

