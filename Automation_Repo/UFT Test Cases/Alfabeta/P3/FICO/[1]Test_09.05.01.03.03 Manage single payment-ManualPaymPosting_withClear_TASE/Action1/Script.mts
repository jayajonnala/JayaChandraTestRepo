		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.05.01.03.03 Manage single payment-ManualPaymPosting_withClear
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
	GetRowNo= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_09.05.01.03.03 Manage single payment-ManualPaymPosting_withClear"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario


'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''--------TransactionCode-F-53 ----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F53_0103_COMPANY_CODE,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F53_0103_AMOUNT,False)
Call SetTextbox("Type","BKPF-BLART","",DT_F53_0103_TYPE,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F53_0103_DOCUMENT_DATE),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F53_0103_CURRENCYRATE,False)
Call SetTextbox("Account","RF05A-KONTO","",DT_F53_0103_ACCOUNT,False)
Call SetTextbox("Account","RF05A-AGKON","",DT_F53_0103_ACCOUNT_OCC1,False)
Call TakeScreenShot

Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
Call TakeScreenShot
Call ClickButton("Sort in Ascending Order",False)
Call TakeScreenShot
Call SetTextbox("Amount","RF05A-SRTFO","",DT_F53_0103_AMOUNT,True)
Call TakeScreenShot
Call ClickButton("Sort Ascending   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Select All",False)
Call TakeScreenShot
Call ClickButton("Deactivate Items",False)
Call TakeScreenShot

Call GetTableCellData("SAPDF05XTC_6102", "Document Number", 1, "", "", "DT_DOCUMENT_NUMBER_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_DOCUMENT_NUMBER_OUTPUT",DT_DOCUMENT_NUMBER)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SelectCellGuiTable("SAPDF05XTC_6102","EUR Gross","Document Number",DT_DOCUMENT_NUMBER, False)
Call Sendkey("{F2}")
Wait(3)
Call TakeScreenShot

Call GetTextboxValue("RF05A-NETTO", 0, "DT_F53_6102_CHECK_TEXT_OF_ASSIGNED_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_F53_6102_CHECK_TEXT_OF_ASSIGNED_OUTPUT",DT_F53_6102_CHECK_TEXT_OF_ASSIGNED)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call ClickButton("Document Overview   \(Shift\+F2\)",False)
Call TakeScreenShot
Call FocusTextBoxByIndex("PK  BusA Acct                               EUR   Amount        Tax amnt","RF05A-AZEI1",0,False)
Call Sendkey("{F2}")
Call Sendkey("{F2}")  ''Added as part of oneclickmaintenance on 21/06/2024

Wait(3)
Call TakeScreenShot
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F53_0300_AMOUNT,False)
Call TakeScreenShot
Call Clickbutton("Display Additional Data for Document Item   \(F7\)",False)
Call TakeScreenShot
Call SetTextbox("House Bank","BSEG-HBKID","",DT_F53_0330_HOUSE_BANK,False)
Call SetTextbox("/","BSEG-HKTID","",DT_F53_0330_BSEGHKTID,False)
Call TakeScreenShot
Call SelectMenuBar("Document;Simulate")
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)  ''Added as part of oneclickmaintenance on 21/06/2024
Call TakeScreenShot

Call GetStatusBar("item1", "DT_F53_0103_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_F53_0103_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_F53_0103_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SelectMenuBar("Document;Display")
Call TakeScreenShot

Call SelectMenuIdToolBar("&COL0",1)
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_F53_0750_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",DT_F53_0750_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)
Call ClickButtonIfExist("Transfer   \(Enter\)",True)

Call VerifyTextBoxContent("Document Number","BKPF-BELNR", 0, DT_F53_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER, False)
Call VerifyTextBoxContent("Company Code","BKPF-BUKRS", 0, DT_F53_0750_CHECK_TEXT_OF_COMPANY_CODE, False)
Call VerifyTextBoxContent("Document Date","BKPF-BLDAT", 0, ConvertDate(DT_F53_0750_CHECK_TEXT_OF_DOCUMENT_DATE), False)
Call VerifyTextBoxContent("Posting Date","BKPF-BUDAT", 0, ConvertDate(DT_F53_0750_CHECK_TEXT_OF_POSTING_DATE), False)

Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_F53_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_F53_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 1, "Assignment", 0, DT_F53_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)
Call VerifyGridCellContent("", 1, "Amount", 0, DT_F53_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("", 1, "Currency", 0, DT_F53_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_RF05A_UBAZW)
Call VerifyGridCellContent("", 1, "G/L Account", 0, DT_F53_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)

DT_F53_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR_OCC1 = year(DT_F53_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR)& ConvertDoubledigit(CSTR(Month(DT_F53_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR)))& ConvertDoubledigit(CSTR(Day(DT_F53_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR)))
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_F53_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_F53_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 2, "Assignment", 0, DT_F53_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR_OCC1)
Call VerifyGridCellContent("", 2, "Amount", 0, DT_F53_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("", 2, "Currency", 0, DT_F53_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_RF05A_UBAZW)
Call VerifyGridCellContent("", 2, "G/L Account", 0, DT_F53_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)

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


