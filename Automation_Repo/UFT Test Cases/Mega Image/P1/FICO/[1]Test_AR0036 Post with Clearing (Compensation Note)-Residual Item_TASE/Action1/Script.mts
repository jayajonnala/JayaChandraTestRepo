
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_AR0036 Post with Clearing (Compensation Note)-Residual Item_TASE
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


gstrTestCaseName = "Test_AR0036 Post with Clearing (Compensation Note)-Residual Item_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


''''----------------------Tcode ZFIGL_UPLOAD_POST----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

'Enter the Details
Call SetTextbox("File path name","P_FILE","",DT_ZFIGL_UPLOAD_POST_1000_FILE_PATH_NAME,False)
Call SetTextbox("Session","P_SESS","",DT_ZFIGL_UPLOAD_POST_1000_SESSION,False)
'Capture the screenshot
Call TakeScreenShot()
'Click execute
Call ClickButton("Execute   \(F8\)",False)
wait 5
'Capture the screenshot
Call TakeScreenShot()
call ClickButton("Execute   \(F8\)",False)

'Capture the screenshot
Call TakeScreenShot()
call ClickButtonIfExist("Yes",True)
wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)
'
'''----------------------Tcode SM35----------------------------
'
'Enter the Tcode
Call SetTcode(DT_ZFIGL_UPLOAD_POST_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
'Capture the screenshot
Call TakeScreenShot()
'select row
Call SetTextbox("Sess\.:","D0100-MAPN","",DT_ZFIGL_UPLOAD_POST_1000_SESSION,False)
Call PressEnter()  
Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI",1,False)
Wait(1)
Call ClickButton("Process session   \(F8\)",False)

'Capture the screenshot
Call TakeScreenShot()
Call SelectRadioButtonIfPopupExists("D0300-BATCH",DT_ZFIGL_UPLOAD_POST_0300_BACKGROUND)
Call ClickButtonIfExist("Process   \(Enter\)",True)
'Capture the screenshot
call TakeScreenShot()

Call ClickButtonIfExist("Go back to batch input session overview   \(Enter\)",True)
'Capture the screenshot
call TakeScreenShot()
Wait(30)
Call SetTextbox("Sess\.:","D0100-MAPN","",DT_ZFIGL_UPLOAD_POST_1000_SESSION,False)
Call PressEnter()  
Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI",1,False)
'Capture the screenshot
call TakeScreenShot()

Call ClickButton("Analyze session   \(F2\)",False)
'Navigate to Log Created Tab
Call SelectTab("TAB_DYNPRO"," Log created on "&ConvertDate(DT_DATE),False)
'Capture the screenshot
call TakeScreenShot()


'fetch oupput mesage
'Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",2,"Index","1","DT_ZFIGL_UPLOAD_POST_0300_MSG1_OUTPUT",False)
'Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",4,"Index","2","DT_ZFIGL_UPLOAD_POST_0300_MSG2_OUTPUT",False)
'Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",6,"Index","3","DT_ZFIGL_UPLOAD_POST_0300_MSG3_OUTPUT",False)
'Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",8,"Index","4","DT_ZFIGL_UPLOAD_POST_0300_MSG4_OUTPUT",False)
'Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",10,"Index","5","DT_ZFIGL_UPLOAD_POST_0300_MSG5_OUTPUT",False)
'Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",12,"Index","6","DT_ZFIGL_UPLOAD_POST_0300_MSG6_OUTPUT",False)
'Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",14,"Index","7","DT_ZFIGL_UPLOAD_POST_0300_MSG7_OUTPUT",False)
'Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",16,"Index","8","DT_ZFIGL_UPLOAD_POST_0300_MSG8_OUTPUT",False)
'Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",18,"Index","9","DT_ZFIGL_UPLOAD_POST_0300_MSG9_OUTPUT",False)
'Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",20,"Index","10","DT_ZFIGL_UPLOAD_POST_0300_MSG10_OUTPUT",False)
'Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",22,"Index","11","DT_ZFIGL_UPLOAD_POST_0300_MSG11_OUTPUT",False)
'Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",24,"Index","12","DT_ZFIGL_UPLOAD_POST_0300_MSG12_OUTPUT",False)
'Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",26,"Index","13","DT_ZFIGL_UPLOAD_POST_0300_MSG13_OUTPUT",False)
'Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",28,"Index","14","DT_ZFIGL_UPLOAD_POST_0300_MSG14_OUTPUT",False)
'Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",30,"Index","15","DT_ZFIGL_UPLOAD_POST_0300_MSG15_OUTPUT",False)
'Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",32,"Index","16","DT_ZFIGL_UPLOAD_POST_0300_MSG16_OUTPUT",False)
'Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",34,"Index","17","DT_ZFIGL_UPLOAD_POST_0300_MSG17_OUTPUT",False)
'Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",36,"Index","18","DT_ZFIGL_UPLOAD_POST_0300_MSG18_OUTPUT",False)
'Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",38,"Index","19","DT_ZFIGL_UPLOAD_POST_0300_MSG19_OUTPUT",False)
'Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",40,"Index","20","DT_ZFIGL_UPLOAD_POST_0300_MSG20_OUTPUT",False)
'Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",43,"Index","21","DT_ZFIGL_UPLOAD_POST_0300_MSG21_OUTPUT",False)
''Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",44,"Index","21","DT_ZFIGL_UPLOAD_POST_0300_MSG21_OUTPUT",False)

Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",3,"Index","1","DT_ZFIGL_UPLOAD_POST_0300_MSG1_OUTPUT",False)
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",6,"Index","2","DT_ZFIGL_UPLOAD_POST_0300_MSG2_OUTPUT",False)
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",9,"Index","3","DT_ZFIGL_UPLOAD_POST_0300_MSG3_OUTPUT",False)
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",12,"Index","4","DT_ZFIGL_UPLOAD_POST_0300_MSG4_OUTPUT",False)
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",15,"Index","5","DT_ZFIGL_UPLOAD_POST_0300_MSG5_OUTPUT",False)
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",18,"Index","6","DT_ZFIGL_UPLOAD_POST_0300_MSG6_OUTPUT",False)
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",21,"Index","7","DT_ZFIGL_UPLOAD_POST_0300_MSG7_OUTPUT",False)
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",24,"Index","8","DT_ZFIGL_UPLOAD_POST_0300_MSG8_OUTPUT",False)
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",27,"Index","9","DT_ZFIGL_UPLOAD_POST_0300_MSG9_OUTPUT",False)
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",30,"Index","10","DT_ZFIGL_UPLOAD_POST_0300_MSG10_OUTPUT",False)
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",33,"Index","11","DT_ZFIGL_UPLOAD_POST_0300_MSG11_OUTPUT",False)
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",36,"Index","12","DT_ZFIGL_UPLOAD_POST_0300_MSG12_OUTPUT",False)
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",39,"Index","13","DT_ZFIGL_UPLOAD_POST_0300_MSG13_OUTPUT",False)
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",42,"Index","14","DT_ZFIGL_UPLOAD_POST_0300_MSG14_OUTPUT",False)
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",45,"Index","15","DT_ZFIGL_UPLOAD_POST_0300_MSG15_OUTPUT",False)
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",48,"Index","16","DT_ZFIGL_UPLOAD_POST_0300_MSG16_OUTPUT",False)
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",51,"Index","17","DT_ZFIGL_UPLOAD_POST_0300_MSG17_OUTPUT",False)
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",54,"Index","18","DT_ZFIGL_UPLOAD_POST_0300_MSG18_OUTPUT",False)
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",57,"Index","19","DT_ZFIGL_UPLOAD_POST_0300_MSG19_OUTPUT",False)
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",60,"Index","20","DT_ZFIGL_UPLOAD_POST_0300_MSG20_OUTPUT",False)
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",65,"Index","21","DT_ZFIGL_UPLOAD_POST_0300_MSG21_OUTPUT",False)

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call VerifyTableCellContentByRef("RSBDC_ANALYSETC_PROTOCOL","No.",364,"Message",Lcase(DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_17))
Call VerifyTableCellContentByRef("RSBDC_ANALYSETC_PROTOCOL","No.",365,"Message",Lcase(DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_18))
Call VerifyTableCellContentByRef("RSBDC_ANALYSETC_PROTOCOL","No.",382,"Message",Lcase(DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_20))

Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)
'
''----------------------Tcode F-30----------------------------
'
'Enter the Tcode
Call SetTcode(DT_ZFIGL_UPLOAD_POST_0100_OKCD_OCC1) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_ZFIGL_UPLOAD_POST_0122_DOCUMENT_DATE),False)
Call SetTextbox("Type","BKPF-BLART","",DT_ZFIGL_UPLOAD_POST_0122_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_ZFIGL_UPLOAD_POST_0122_COMPANY_CODE,False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_ZFIGL_UPLOAD_POST_0122_POSTING_DATE),False)
'Call SetTextbox("Period","BKPF-MONAT","",DT_F30_0122_PERIOD,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_ZFIGL_UPLOAD_POST_0122_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_ZFIGL_UPLOAD_POST_0122_REFERENCE,False)
'Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F30_0122_DOCHEADER_TEXT,False)

'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Choose open items   \(F6\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Company Code","RF05A-AGBUK","",DT_ZFIGL_UPLOAD_POST_0710_COMPANY_CODE,False)
Call SetTextbox("Account","RF05A-AGKON","",DT_ZFIGL_UPLOAD_POST_0710_ACCOUNT,False)
'Call SetTextbox("Account Type","RF05A-AGKOA","",DT_F30_0710_ACCOUNT_TYPE,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Select All",False)
Call ClickButton("Deactivate Items",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Field content search",False)
Call SelectRadioButton("RF05A-XPOS1",DT_ZFIGL_UPLOAD_POST_2000_DOCUMENT_NUMBER,True)
Call ClickButton("Continue   \(Enter\)",True)

Call SetTextbox("From","RF05A-SEL01","0",DT_ZFIGL_UPLOAD_POST_0731_FROM,True)
Call SetTextbox("To","RF05A-SEL02","0",DT_ZFIGL_UPLOAD_POST_0731_TO,True)
Call SetTextbox("From","RF05A-SEL01","1",DT_ZFIGL_UPLOAD_POST_0731_FROM_OCC1,True)
Call SetTextbox("To","RF05A-SEL02","1",DT_ZFIGL_UPLOAD_POST_0731_TO_OCC1,True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)

Call ClickButton("Select All",False)
Call ClickButton("Activate Items",False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TS","Partial Pmt",False)
'Call GetTableCellData("SAPDF05XTC_6104","Payment Amount",6,"","","DT_ROW_6_OUTPUT",False)
'Call GetTextboxValue("RF05A-DIFFB","","DT_ASSINGED_OUTPUT",False)
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'' FindRowNumber(tableNameOrGridTitle, refColumnName, refCellValue, dataTableColumnName)
Call FindRowNumber("SAPDF05XTC_6104","Document Number",DT_ZFIGL_UPLOAD_POST_0731_TO,"Row_Number_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("Row_Number_OUTPUT",Row_Number)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'Call SetTableDataNoRef("SAPDF05XTC_6104","Payment Amount",6,DT_ZFIGL_UPLOAD_POST_6104_TABLECELL_PAYMENT_AMOUNT_5_OCC1,False)
Call SetTableDataNoRef("SAPDF05XTC_6104","Payment Amount",Row_Number,DT_ZFIGL_UPLOAD_POST_6104_TABLECELL_PAYMENT_AMOUNT_5_OCC2,False)
Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call VerifyTextBoxContent("Not assigned","RF05A-DIFFB","",DT_ASSINGED_AMOUNT,False)
Call VerifyTextBoxContent("Amount entered","RF05A-BETRG","",DT_ASSINGED_AMOUNT,False)

Call SelectMenuBar("Document;Simulate")
Wait(2)
Call VerifyTextBoxContent("C","RF05A-AZSAL","",DT_ASSINGED_AMOUNT,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
''Validate If invoice is generated
Call GetStatusBar("item1","DT_DOC_NO_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call VerifyStatusBarMessageType("S")
VerifyStatusBar(DT_ZFIGL_UPLOAD_POST_0122_CHECK_TEXT_OF_STATUSBAR)

Call ClickButton("Back   \(F3\)",False)
Call ClickButtonIfExist("Yes",True)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)

''----------------------Tcode FB03----------------------------
'
'Enter the Tcode
Call SetTcode(DT_ZFIGL_UPLOAD_POST_0100_OKCD_OCC2) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC6)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Number","RF05L-BELNR","",DT_ZFIGL_UPLOAD_POST_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_ZFIGL_UPLOAD_POST_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_ZFIGL_UPLOAD_POST_0100_FISCAL_YEAR,False)

Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

Call VerifyGridCellContent("",1,"BSCHL","",DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",2,"BSCHL","",DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("",3,"BSCHL","",DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("",5,"BSCHL","",DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BSCHL)
Call VerifyGridCellContent("",4,"BSCHL","",DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BSCHL)

Call VerifyGridCellContent("",1,"Account","",DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("",2,"Account","",DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("",3,"Account","",DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
Call VerifyGridCellContent("",5,"Account","",DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_KTONR)
Call VerifyGridCellContent("",4,"Account","",DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KTONR)

Call VerifyGridCellContent("",1,"Amount","",DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMBTR)
Call VerifyGridCellContent("",2,"Amount","",DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMBTR)
Call VerifyGridCellContent("",3,"Amount","",DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_DMBTR)
Call VerifyGridCellContent("",4,"Amount","",DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_DMBTR)
Call VerifyGridCellContent("",5,"Amount","",DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_DMBTR)

Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)
''
'''''----------------------Tcode FBL1N----------------------------
'''Enter the Tcode
Call SetTcode(DT_ZFIGL_UPLOAD_POST_0100_OKCD_OCC3) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC8)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Company code","KD_BUKRS-LOW","",DT_ZFIGL_UPLOAD_POST_1000_COMPANY_CODE,False)
Call SetTextbox("Vendor account","KD_LIFNR-LOW","",DT_ZFIGL_UPLOAD_POST_1000_VENDOR_ACCOUNT,False)

'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Dynamic selections   \(Shift\+F4\)",False)

Call ActivateNodeGuiTree("","Vendor master;Industry")
Call ActivateNodeGuiTree("","Vendor master;Group key")
Call ActivateNodeGuiTree("","Vendor master;Country")
Call ActivateNodeGuiTree("","Vendor master;Postal Code")
Call ActivateNodeGuiTree("","Vendor master;City")
Call ActivateNodeGuiTree("","Vendor master;Trading partner")
Call ActivateNodeGuiTree("","Document;Special G/L ind.")
Call ActivateNodeGuiTree("","Document;Assignment")
Call ActivateNodeGuiTree("","Company Code;Clerk Abbreviation")
Call ActivateNodeGuiTree("","Company Code;Reconciliation acct")
'Capture the screenshot
Call TakeScreenShot()

wait(1)
SapGuiSession("transaction:=FBL1N").SapGuiWindow("transaction:=FBL1N").SAPGuiButton("tooltip:=Multiple selection","index:=5").Click

'Capture the screenshot
Call TakeScreenShot()

'set filter criteria
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",1,DT_ZFIGL_UPLOAD_POST_1105_CLEARING_DOCUMENT,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",2,DT_ZFIGL_UPLOAD_POST_0731_FROM_OCC1,True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Copy   \(F8\)",True)

Call SelectRadioButton("X_AISEL","All items",FAlse)
Call SetTextbox("Posting date","SO_BUDAT-LOW","",DT_ZFIGL_UPLOAD_POST_1000_POSTING_DATE,False)
Call SetTextbox("to","SO_BUDAT-HIGH","",DT_ZFIGL_UPLOAD_POST_1000_TO,False)

Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Call VerifyifGuiLabelWithIndexExists(DT_ZFIGL_UPLOAD_POST_0120_CHECK_TEXT_OF_NO_NAME,"1")
Call VerifyifGuiLabelExists(DT_ZFIGL_UPLOAD_POST_0120_CHECK_TEXT_OF_ZO)
''Call VerifyifGuiLabelExists(DT_ZFIGL_UPLOAD_POST_0120_CHECK_TEXT_OF_NO_NAME_OCC1)
Call VerifyifGuiLabelWithIndexExists(DT_ZFIGL_UPLOAD_POST_0120_CHECK_TEXT_OF_NO_NAME_OCC1,"1")
Call VerifyifGuiLabelExists(DT_ZFIGL_UPLOAD_POST_0120_CHECK_TEXT_OF_NO_NAME_OCC3)
Call VerifyifGuiLabelExists(DT_ZFIGL_UPLOAD_POST_0120_CHECK_TEXT_OF_KR)
''Call VerifyifGuiLabelExists(DT_ZFIGL_UPLOAD_POST_0120_CHECK_TEXT_OF_NO_NAME_OCC4)
Call VerifyifGuiLabelWithIndexExists(DT_ZFIGL_UPLOAD_POST_0120_CHECK_TEXT_OF_NO_NAME_OCC4,"1")

Call TakeScreenShot()
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)

Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)
''
'''----------------------Tcode FBL5N----------------------------
''Enter the Tcode
Call SetTcode(DT_ZFIGL_UPLOAD_POST_0100_OKCD_OCC4) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC10)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_ZFIGL_UPLOAD_POST_1000_CUSTOMER_ACCOUNT,False)
Call SelectRadioButton("X_AISEL","All items",False)
Call ClickButtonIfExist("Dynamic selections   \(Shift\+F4\)",False)
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_ZFIGL_UPLOAD_POST_1000_COMPANY_CODE_OCC1,False)

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
Call SelectTab("TAB_STRIP","Select Ranges",True)
'set filter criteria
Call SetTableDataNoRef("SAPLALDBINTERVAL","Lower limit",1,DT_ZFIGL_UPLOAD_POST_0731_FROM,True)
Call SetTableDataNoRef("SAPLALDBINTERVAL","Upper limit",1,DT_ZFIGL_UPLOAD_POST_0731_TO,True)
Call SetTableDataNoRef("SAPLALDBINTERVAL","Lower limit",2,DT_ZFIGL_UPLOAD_POST_0731_FROM_OCC1,True)
Call SetTableDataNoRef("SAPLALDBINTERVAL","Upper limit",2,DT_ZFIGL_UPLOAD_POST_0731_TO_OCC1,True)
Call SetTableDataNoRef("SAPLALDBINTERVAL","Lower limit",3,DT_DOC_NO,True)
Call SetTableDataNoRef("SAPLALDBINTERVAL","Upper limit",3,DT_DOC_NO,True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Copy   \(F8\)",True)

'Capture the screenshot
Call TakeScreenShot()
'
Call ClickButton("Execute   \(F8\)",False)
Wait(2)
Call TakeScreenShot()
Call VerifyStatusBarMessageType("S")

Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_1105_DOCUMENT_NUMBER,0)
Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_0120_CHECK_TEXT_OF_NO_NAME_OCC10,1)
Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_0120_CHECK_TEXT_OF_NO_NAME_OCC8,0)
Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_0120_CHECK_TEXT_OF_NO_NAME_OCC11,0)
Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_1105_DOCUMENT_NUMBER_OCC1,0)
Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_0120_CHECK_TEXT_OF_NO_NAME_OCC14,1)

Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_0300_MSG1,0)
Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_0300_MSG2,0)
Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_0300_MSG3,0)
Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_0300_MSG4,0)
Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_0300_MSG5,0)
Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_0300_MSG6,0)
Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_0300_MSG7,0)
Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_0300_MSG8,0)
Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_0300_MSG9,0)
Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_0300_MSG10,0)
Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_0300_MSG11,0)
Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_0300_MSG12,0)
Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_0300_MSG13,0)
Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_0300_MSG14,0)
Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_0300_MSG15,0)
Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_0300_MSG16,0)
Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_0300_MSG17,0)

' SetHorizontalScrollBar(ScrollBarPosition, blnIsItPopup)
Call SetVerticalScrollBar(20,False)
Wait(10)

Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_0300_MSG18,0)
Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_0300_MSG19,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_0300_MSG20,0)

Wait(10)
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)
'
'''----------------------Tcode FB12----------------------------
'
'Enter the Tcode
Call SetTcode(DT_ZFIGL_UPLOAD_POST_0100_OKCD_OCC5) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC12)
'Capture the screenshot
'Call TakeScreenShot()

Call SetTextbox("Company Code","RF022-BUKRS","",DT_ZFIGL_UPLOAD_POST_0100_COMPANY_CODE_OCC1,False)
'Capture the screenshot
'Call TakeScreenShot()

Call PressEnter() 
'Capture the screenshot
'Call TakeScreenShot()
Call ClickButton("Find   \(Ctrl\+F\)",True)
Call SetTextbox("Find","RSYSF-STRING","",DT_ZFIGL_UPLOAD_POST_0800_FIND,True)
Call ClickButton("Find   \(Enter\)",True)
Call ClickLabel(DT_ZFIGL_UPLOAD_POST_0800_FIND,"",True)
Call ClickButton("Copy   \(Enter\)",True)

Call SetTextbox("Document Number","RF022-BELNR","",DT_ZFIGL_UPLOAD_POST_1001_DOCUMENT_NUMBER,True)
Call SetTextbox("Fiscal Year","RF022-GJAHR","",DT_ZFIGL_UPLOAD_POST_1001_FISCAL_YEAR,True)
'Capture the screenshot
'Call TakeScreenShot()

Call ClickButton("Continue   \(Enter\)",True)
Wait(1)
'Capture the screenshot
'Call TakeScreenShot()
'verify status bar
Call VerifyStatusBar(DT_ZFIGL_UPLOAD_POST_0100_CHECK_TEXT_OF_STATUSBAR)

'Call TakeScreenShot()
Call ClickButton("Back   \(F3\)",False)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)
'
'''----------------------Tcode F.64----------------------------
'
'Enter the Tcode
Call SetTcode(DT_ZFIGL_UPLOAD_POST_0100_OKCD_OCC6) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC14)
'Capture the screenshot
'Call TakeScreenShot()

Call SetTextbox("Correspondence","EVENT-LOW","",DT_ZFIGL_UPLOAD_POST_1000_CORRESPONDENCE,False)
Call SetTextbox("Company code","BUKRS-LOW","",DT_ZFIGL_UPLOAD_POST_1000_COMPANY_CODE_OCC2,False)
Call SetTextbox("Document number","BELNR-LOW","",DT_ZFIGL_UPLOAD_POST_1000_DOCUMENT_NUMBER,False)
Call SetTextbox("Fiscal year","GJAHR-LOW","",DT_ZFIGL_UPLOAD_POST_1000_FISCAL_YEAR,False)
'Capture the screenshot
'Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Wait(1)
'Capture the screenshot
'Call TakeScreenShot()

Call VerifyifGuiLabelExists(DT_ZFIGL_UPLOAD_POST_0120_CHECK_TEXT_OF_NO_NAME_OCC36)
Call ClickLabel(DT_ZFIGL_UPLOAD_POST_0120_CHECK_TEXT_OF_NO_NAME_OCC36,"",False)

Call SetTextboxPopupIfExist("USR01-SPLD","Output Device",DT_ZFIGL_UPLOAD_POST_1100_OUTPUT_DEVICE)
'Capture the screenshot
'Call TakeScreenShot()

Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(3)
'Capture the screenshot
'Call TakeScreenShot()

Call SetTextboxPopupIfExist("USR01-SPLD","Output Device",DT_ZFIGL_UPLOAD_POST_1100_OUTPUT_DEVICE)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(4)
'Capture the screenshot
'Call TakeScreenShot()

Wait(2)
'Capture the screenshot
'Call TakeScreenShot()


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


