
'''''''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'''''''.................Test Script Name :Test_09.05.01.16.01 Post In_Outgoing Payment - Manual Incoming Paymen
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

gstrTestCaseName = "Test_09.05.01.16.01 Post In_Outgoing Payment - Manual Incoming Paymen"
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

'''''--------TransactionCode-FB01----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB01_0100_COMPANY_CODE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_FB01_0100_TYPE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB01_0100_PSTKY,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_FB01_0100_DOCUMENT_DATE),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FB01_0100_CURRENCYRATE,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB01_0100_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB01_0302_AMOUNT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB01_0302_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","1",DT_FB01_0302_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB01_0300_AMOUNT,False)
Call TakeScreenShot
Call ClickButton("Display Additional Data for Document Item   \(F7\)",False)
Call TakeScreenShot
Call SetTextbox("House Bank","BSEG-HBKID","",DT_FB01_0330_HOUSE_BANK,False)
Call SetTextbox("/","BSEG-HKTID","",DT_FB01_0330_BSEGHKTID,False)
Call TakeScreenShot
Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
Call TakeScreenShot

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)

Call GetStatusBar("item1","DT_FB01_0100_GET_DOCUMENTNUMBER_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FB01_0100_GET_DOCUMENTNUMBER_OUTPUT",DT_FB01_0100_GET_DOCUMENTNUMBER)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyStatusBar(Lcase(DT_FB01_0100_CHECK_TEXT_OF_STATUSBAR))

''''--------TransactionCode-F-28----------''''
Call SetTcode(DT_FB01_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB01_0103_COMPANY_CODE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_FB01_0103_TYPE,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_FB01_0103_DOCUMENT_DATE),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FB01_0103_CURRENCYRATE,False)
Call SetTextbox("Account","RF05A-KONTO","",DT_FB01_0103_ACCOUNT,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB01_0103_AMOUNT,False)
Call SetTextbox("Value date","BSEG-VALUT","",ConvertDate(DT_FB01_0103_VALUE_DATE),False)
Call SetTextbox("Account","RF05A-AGKON","",DT_FB01_0103_ACCOUNT_OCC1,False)
Call SetTextbox("Account Type","RF05A-AGKOA","",DT_FB01_0103_ACCOUNT_TYPE,False)
Call TakeScreenShot
Call Clickbutton("Process Open Items   \(Shift\+F4\)",False)
Call TakeScreenShot

Call ClickButton("Deselect All",False)
Call TakeScreenShot
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC2)
Call ClickButton("Select All",False)
Call TakeScreenShot
Call ClickButton("Deactivate Items",False)
Call TakeScreenShot
Call SelectTab("TS", "Partial Pmt", False)
Call TakeScreenShot
Call ClickButton("Field content search",False)
Call TakeScreenShot

Call SelectRadioButton("RF05A-XPOS1", "Document Number", True)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("From","RF05A-SEL01","",DT_FB01_0731_FROM,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot
'Call FocusTableCell("SAPDF05XTC_6104",  "Net Amount", "", "Document Number", DT_FB01_0731_FROM, "", False)
Call SelectCellGuiTable("SAPDF05XTC_6104", "Net Amount", "Document Number", DT_FB01_0731_FROM, False)
Call SendKey("{F2}")
Wait 2
Call SetTableData("SAPDF05XTC_6104", "Payment Amount", "", "Document Number", DT_FB01_0731_FROM, DT_FB01_6104_TABLECELL_PAYMENT_AMOUNT_0, False)
Call TakeScreenShot
Call ClickButton("Document Overview   \(Shift\+F2\)",False)
Call TakeScreenShot
Call FocusTextBoxByIndex("PK  BusA Acct                               EUR   Amount        Tax amnt","RF05A-AZEI1",0,False)
Call Sendkey("{F2}")
Wait(3)
Call TakeScreenShot
Call Clickbutton("Display Additional Data for Document Item   \(F7\)",False)
Call TakeScreenShot
Call SetTextbox("House Bank","BSEG-HBKID","",DT_FB01_0330_HOUSE_BANK_OCC1,False)
Call SetTextbox("/","BSEG-HKTID","",DT_FB01_0330_BSEGHKTID_OCC1,False)
Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
Call TakeScreenShot
Call SelectMenuBar("Document;Simulate")
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call VerifyStatusBarMessageType("S")
Call GetStatusBar("item1", "DT_FB01_0103_GET_TEXT_OF_STATUSBAR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_FB01_0103_GET_TEXT_OF_STATUSBAR_OUTPUT",DT_FB01_0103_GET_TEXT_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

'''''--------TransactionCode-FB03 ----------''''
Call SetTcode(DT_FB01_0100_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB01_0100_DOCUMENT_NUMBER	,False) 
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB01_0100_COMPANY_CODE_OCC1,False) 
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",Year(DT_FB01_0100_FISCAL_YEAR),False) 
Call PressEnter()     
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "KOBEZ", 0, DT_FB01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOBEZ)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_FB01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 1, "Amount", 0, DT_FB01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)

Call VerifyGridCellContent("", 2, "KOBEZ", 0, DT_FB01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KOBEZ)
Call VerifyGridCellContent("", 2, "AZBET", 0, DT_FB01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_FB01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)


Call LogOff'









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





