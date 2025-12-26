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
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Manage Petty Cash Transactions_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 3rd May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Manage Petty Cash Transactions_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Manage Petty Cash Transactions_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
'''----------------------Tcode FBCJ----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

'Enter the Details
'Call SetTextbox("","F_DISPLAY_PERIOD_LO","",Replace(DT_FBCJ_0100_F_DISPLAY_PERIOD_LO,"/","."),False)
'Call SetTextbox("-","F_DISPLAY_PERIOD_HI","",Replace(DT_FBCJ_0100_F_DISPLAY_PERIOD_HI,"/","."),False)
Call ClickButton("FB_TODAY",False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

Call GetTextboxValue("F_NUMB_OF_REC","","DT_LAST_POPULATED_ROW_OUTPUT",False)
Call GetTextboxValue("F_NUMB_OF_PAYM","","DT_GET_ROW_TO_SELECT_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SelectTab("F_TABSTRIP",DT_FBCJ_0100_CASH_RECEIPTS,False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTableDataNoRef("SAPMFCJ0FTCJ_R_POSTINGS","Business transaction",DT_ROW_FIRST_DOCUMENT,DT_FBCJ_0120_TABLECELL_BUSINESS_TRANSACTION_0,False)
Call SetTableDataNoRef("SAPMFCJ0FTCJ_R_POSTINGS","Amount",DT_ROW_FIRST_DOCUMENT,DT_FBCJ_0120_TABLECELL_AMOUNT_0,False)
'''Call SetTableDataNoRef("SAPMFCJ0FTCJ_R_POSTINGS","Vendor",DT_ROW_FIRST_DOCUMENT,DT_FBCJ_0120_TABLECELL_VENDOR_0,False)
Call SetTableDataNoRef("SAPMFCJ0FTCJ_R_POSTINGS","Supplier",DT_ROW_FIRST_DOCUMENT,DT_FBCJ_0120_TABLECELL_VENDOR_0,False)

'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()
Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

Call SetTableDataNoRef("SAPMFCJ0FTCJ_R_POSTINGS","Business transaction",DT_ROW_SECOND_DOCUMENT,DT_FBCJ_0120_TABLECELL_BUSINESS_TRANSACTION_1,False)
Call SetTableDataNoRef("SAPMFCJ0FTCJ_R_POSTINGS","Amount",DT_ROW_SECOND_DOCUMENT,DT_FBCJ_0120_TABLECELL_AMOUNT_1,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

Call SetTableDataNoRef("SAPMFCJ0FTCJ_R_POSTINGS","Business transaction",DT_ROW_THIRD_DOCUMENT,DT_FBCJ_0120_TABLECELL_BUSINESS_TRANSACTION_2,False)
Call SetTableDataNoRef("SAPMFCJ0FTCJ_R_POSTINGS","Amount",DT_ROW_THIRD_DOCUMENT,DT_FBCJ_0120_TABLECELL_AMOUNT_2,False)
'''Call SetTableDataNoRef("SAPMFCJ0FTCJ_R_POSTINGS","Vendor",DT_ROW_THIRD_DOCUMENT,DT_FBCJ_0120_TABLECELL_VENDOR_2,False)
Call SetTableDataNoRef("SAPMFCJ0FTCJ_R_POSTINGS","Supplier",DT_ROW_THIRD_DOCUMENT,DT_FBCJ_0120_TABLECELL_VENDOR_2,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()
Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Post all entries   \(F6\)",False)
Wait(3)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyTableCellContent(DT_ROW_FIRST_DOCUMENT,"Doc.Stat.","SAPMFCJ0FTCJ_R_POSTINGS","S_TL_G")
Call VerifyTableCellContent(DT_ROW_SECOND_DOCUMENT,"Doc.Stat.","SAPMFCJ0FTCJ_R_POSTINGS","S_TL_G")
Call VerifyTableCellContent(DT_ROW_THIRD_DOCUMENT,"Doc.Stat.","SAPMFCJ0FTCJ_R_POSTINGS","S_TL_G")

Call GetTableCellData("SAPMFCJ0FTCJ_R_POSTINGS","Internal doc. number",DT_ROW_FIRST_DOCUMENT,"Amount",DT_FBCJ_0120_TABLECELL_AMOUNT_0,"DT_FIRST_DOCUMENT_NUMBER_OUTPUT",False)
Call GetTableCellData("SAPMFCJ0FTCJ_R_POSTINGS","Internal doc. number",DT_ROW_SECOND_DOCUMENT,"Amount",DT_FBCJ_0120_TABLECELL_AMOUNT_1,"DT_SECOND_DOCUMENT_NUMBER_OUTPUT",False)
Call GetTableCellData("SAPMFCJ0FTCJ_R_POSTINGS","Internal doc. number",DT_ROW_THIRD_DOCUMENT,"Amount",DT_FBCJ_0120_TABLECELL_AMOUNT_2,"DT_THIRD_DOCUMENT_NUMBER_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SelectRowGuiTableByRow("SAPMFCJ0FTCJ_R_POSTINGS",DT_ROW_THIRD_DOCUMENT,False)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Print Receipt",False)
Wait(6)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()
'
Call ClickButton("Print cash journal   \(F7\)",False)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

CAll ClickLabel("Document N","",False)
Call ClickButton("Set filter   \(Ctrl\+F5\)",False)
Call ClickButton("Multiple selection",True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",1,DT_FIRST_DOCUMENT_NUMBER,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",2,DT_SECOND_DOCUMENT_NUMBER,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",3,DT_THIRD_DOCUMENT_NUMBER,True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Copy   \(F8\)",True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Execute   \(Enter\)",True)
'Capture the screenshot
Call TakeScreenShot()

CAll ClickLabel("4","0",False)
CAll ClickLabel("4","1",False)
CAll ClickLabel("4","2",False)
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
					Call ReporterFunction(strLibraryFileName,"VerifyifGuiLabelWithIndexExists","1"," Status bar Content","Gui Label with value "&Content &" index : "&Index &" doesn't exist  in the screen")	
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

'Call VerifyifGuiLabelWithIndexExists(DT_FBCJ_0120_CHECK_TEXT_OF_VARIOUS_INC_R,1)
Call VerifyifGuiLabelExists(DT_FIRST_DOCUMENT_NUMBER)

Call VerifyifGuiLabelWithIndexExists(DT_FBCJ_0120_CHECK_TEXT_OF_PETTY_CASH_WIT,1)

'Call VerifyifGuiLabelWithIndexExists(DT_FBCJ_0120_CHECK_TEXT_OF_VARIOUS_INC_R_OCC1,3)
Call VerifyifGuiLabelExists(DT_THIRD_DOCUMENT_NUMBER)


Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("F_TABSTRIP",DT_FBCJ_0100_CASH_PAYMENTS,False)
'Capture the screenshot
Call TakeScreenShot()
Call SetTableDataNoRef("SAPMFCJ0FTCJ_E_POSTINGS","Business transaction",DT_NO_ROWS_TAB1_FINAL2,DT_FBCJ_0110_TABLECELL_BUSINESS_TRANSACTION_0,False)
Call SetTableDataNoRef("SAPMFCJ0FTCJ_E_POSTINGS","Amount",DT_NO_ROWS_TAB1_FINAL2,DT_FBCJ_0110_TABLECELL_AMOUNT_0,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Post all entries   \(F6\)",False)
Wait(3)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyTableCellContent(DT_NO_ROWS_TAB1_FINAL2,"Doc.Stat.","SAPMFCJ0FTCJ_E_POSTINGS","S_TL_G")
Call GetTableCellData("SAPMFCJ0FTCJ_E_POSTINGS","Internal doc. number",DT_NO_ROWS_TAB1_FINAL2,"Amount",DT_FBCJ_0110_TABLECELL_AMOUNT_0,"DT_FOURTH_DOCUMENT_NUMBER_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SelectRowGuiTableByRow("SAPMFCJ0FTCJ_E_POSTINGS",DT_NO_ROWS_TAB1_FINAL2,False)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Print Receipt",False)
Wait(6)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Print cash journal   \(F7\)",False)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

CAll ClickLabel("Document N","",False)
Call ClickButton("Set filter   \(Ctrl\+F5\)",False)
Call ClickButton("Multiple selection",True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",1,DT_FOURTH_DOCUMENT_NUMBER,True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Copy   \(F8\)",True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Execute   \(Enter\)",True)
'Capture the screenshot
Call TakeScreenShot()

CAll ClickLabel("4","0",False)
'Capture the screenshot
Call TakeScreenShot()
Call VerifyifGuiLabelWithIndexExists(DT_FBCJ_0120_CHECK_TEXT_OF_PETTY_CASH_DEP,1)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()


