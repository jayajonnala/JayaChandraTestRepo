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

'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_P2P_01_01_085-PO VAT dif Invoice VAT Domestic vendor_P4_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 3rd June
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_P2P_01_01_085_P4_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_P2P_01_01_085-PO VAT  dif Invoice VAT  Domestic vendor_P4_TASE.xls"
'''''----------------------Login----------------------------

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

'''----------------------Tcode MIGO----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetComboByKey("GODYNPRO-ACTION",DT_MB90_0010_GODYNPROACTION)
'Capture the screenshot
Call TakeScreenShot()

Call SetComboByKey("GODYNPRO-REFDOC",DT_MIGO_0010_GODYNPROREFDOC)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextboxNoLabel("GODYNPRO-MAT_DOC",0,DT_MB90_1000_ARTICLE_DOCUMENT,False) 'DT_MIGO_2010_GODYNPROMAT_DOC
Call SetTextboxNoLabel("GODYNPRO-DOC_YEAR",0,DT_MIGO_1000_FISCAL_YEAR,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()

Call SelectTab("TS_GOHEAD",DT_MIGO_0100_DOC_INFO,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("FI Documents",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()
If SAPGuisession(sessionObject).sapguiwindow(windowObject).sapguiedit("guicomponenttype:=31","name:=BKPF-BELNR","attachedtext:=Document Number","Index:=0").Exist(1) Then
	wait 1
Else
Call DoubleClickGuiGridCell("Documents in Accounting", 0, 1, "Document Number", True)
'Call DoubleClickGuiGridCell("Documents in Accounting", 0, 1, "Doc. Number", True)
End If


Call GetTextboxValue("BKPF-BELNR","","DT_MIGO_0750_GET_TEXT_OF_DOCUMENT_NUMBER_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call VerifyTextBoxContent("Document Number","BKPF-BELNR","",DT_MIGO_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OCC1,False)
Call VerifyTextBoxContent("Company Code","BKPF-BUKRS","",DT_MIGO_0750_CHECK_TEXT_OF_COMPANY_CODE,False)

'verify the grid coponents
call VerifyGridCellContent("",1,"Posting Key","",DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
call VerifyGridCellContent("",2,"Posting Key","",DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
call VerifyGridCellContent("",11,"Posting Key","",DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
call VerifyGridCellContent("",12,"Posting Key","",DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BSCHL)

call VerifyGridCellContent("",1,"Account","",DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
call VerifyGridCellContent("",2,"Account","",DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
call VerifyGridCellContent("",11,"Account","",DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
call VerifyGridCellContent("",12,"Account","",DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KTONR)

call VerifyGridCellContent("",1,"Profit Center","",DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
call VerifyGridCellContent("",2,"Profit Center","",DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)
call VerifyGridCellContent("",11,"Profit Center","",DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_PRCTR)
call VerifyGridCellContent("",12,"Profit Center","",DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_PRCTR)

'''----------------------Tcode MIRO----------------------------
'Enter the Tcode
Call SetTcode(DT_MIGO_0750_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_MIGO_0750_OKCD)
'Capture the screenshot
Call TakeScreenShot()
Call SetCombo("RM08M-VORGANG","Invoice")
'Call SetTextbox("Invoice date","INVFO-BLDAT","",DT_MIGO_0010_INVOICE_DATE,False)
Call SetTextboxNoLabel("INVFO-BLDAT","",ConvertDate(Date),False)
Call PressEnter()
Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIGO_0010_REFERENCE,False)
Call PressEnter()
Call SetTextboxNoLabel("RM08M-EBELN",0,DT_MIGO_6211_RM08MEBELN,False)

Call PressEnter()
Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

Call SetTableDataNoRef("SAPLMR1MTC_MR1M","Tax Code",1,DT_MIGO_6310_TABLECELL_TAX_CODE_0,False)
Call SetTableDataNoRef("SAPLMR1MTC_MR1M","Tax Code",2,DT_MIGO_6310_TABLECELL_TAX_CODE_1,False)
Call SetTableDataNoRef("SAPLMR1MTC_MR1M","Tax Code",3,DT_MIGO_6310_TABLECELL_TAX_CODE_2,False)
Call SetTableDataNoRef("SAPLMR1MTC_MR1M","Tax Code",4,DT_MIGO_6310_TABLECELL_TAX_CODE_3,False)
Call PressEnter()
Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

'Select Calculate Tax field as true
Call SelectCheckbox("INVFO-XMWST",0,DT_MIGO_0010_CALCULATE_TAX,False)
'Call SetComboByKey("RM08M-VORGANG",1)
Call TakeScreenShot()

'Get the remaining balance and enter it in Amount Field
Call GetTextboxValue("RM08M-DIFFERENZ",0,"DT_MIGO_6000_GET_TEXT_OF_BALANCE_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTextbox("Amount","INVFO-WRBTR","",DT_MIGO_0010_AMOUNT,False)
Call PressEnter()
Call PressEnter()
Wait(1)
Call TakeScreenShot()

Call ClickButtonIfExist("Simulate Document   \(Ctrl\+Shift\+F7\)",False)
Wait(1)
Call TakeScreenShot()

Call VerifyTableCellContentPopup(1,"G/L","SAPLMR1MTC_MR1M_PB",DT_MIGO_6250_CHECK_TEXT_OF_TABLECELL_GL_0)
Call VerifyTableCellContentPopup(2,"G/L","SAPLMR1MTC_MR1M_PB",DT_MIGO_6250_CHECK_TEXT_OF_TABLECELL_GL_1)
Call VerifyTableCellContentPopup(3,"G/L","SAPLMR1MTC_MR1M_PB",DT_MIGO_6250_CHECK_TEXT_OF_TABLECELL_GL_2)
Call VerifyTableCellContentPopup(4,"G/L","SAPLMR1MTC_MR1M_PB",DT_MIGO_6250_CHECK_TEXT_OF_TABLECELL_GL_3)
Call VerifyTableCellContentPopup(5,"G/L","SAPLMR1MTC_MR1M_PB",DT_MIGO_6250_CHECK_TEXT_OF_TABLECELL_GL_4)
Call VerifyTableCellContentPopup(6,"G/L","SAPLMR1MTC_MR1M_PB",DT_MIGO_6250_CHECK_TEXT_OF_TABLECELL_GL_5)

Call VerifyTableCellContentPopup(2,"Article","SAPLMR1MTC_MR1M_PB",DT_MIGO_6250_CHECK_TEXT_OF_TABLECELL_ARTICLE_1)
Call VerifyTableCellContentPopup(3,"Article","SAPLMR1MTC_MR1M_PB",DT_MIGO_6250_CHECK_TEXT_OF_TABLECELL_ARTICLE_2)
Call VerifyTableCellContentPopup(4,"Article","SAPLMR1MTC_MR1M_PB",DT_MIGO_6250_CHECK_TEXT_OF_TABLECELL_ARTICLE_3)
Call VerifyTableCellContentPopup(5,"Article","SAPLMR1MTC_MR1M_PB",DT_MIGO_6250_CHECK_TEXT_OF_TABLECELL_ARTICLE_4)

Call VerifyTableCellContentPopup(2,"Profit Center","SAPLMR1MTC_MR1M_PB",DT_MIGO_6250_CHECK_TEXT_OF_TABLECELL_PROFIT_CENTER_1)
Call VerifyTableCellContentPopup(3,"Profit Center","SAPLMR1MTC_MR1M_PB",DT_MIGO_6250_CHECK_TEXT_OF_TABLECELL_PROFIT_CENTER_2)
Call VerifyTableCellContentPopup(4,"Profit Center","SAPLMR1MTC_MR1M_PB",DT_MIGO_6250_CHECK_TEXT_OF_TABLECELL_PROFIT_CENTER_3)
Call VerifyTableCellContentPopup(5,"Profit Center","SAPLMR1MTC_MR1M_PB",DT_MIGO_6250_CHECK_TEXT_OF_TABLECELL_PROFIT_CENTER_4)


'Click on Post Buton
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",True)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
'Validate If invoice is generated
Call GetStatusBar("item1","DT_ACCOUNTING_DOC_NO_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
wait(1)
VerifyStatusBar(DT_MIGO_6000_CHECK_TEXT_OF_STATUSBAR_OCC1)
Call VerifyStatusBarMessageType("S")


Call SelectMenuBar("Invoice Document;Display")
Wait(1)
Call TakeScreenShot()
Call ClickButtonIfExist("Follow-On Documents \.\.\.   \(F8\)",False)
Call TakeScreenShot()

Call GetTextboxValue("BKPF-BELNR","","DT_MIGO_0750_GET_TEXT_OF_DOCUMENT_NUMBER_OCC2_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'verify the grid coponents
call VerifyGridCellContent("",1,"Posting Key","",DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OCC1)
call VerifyGridCellContent("",2,"Posting Key","",DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL_OCC1)
call VerifyGridCellContent("",3,"Posting Key","",DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL_OCC1)
call VerifyGridCellContent("",4,"Posting Key","",DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BSCHL_OCC1)

call VerifyGridCellContent("",1,"Account","",DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR_OCC1)
call VerifyGridCellContent("",2,"Account","",DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR_OCC1)
call VerifyGridCellContent("",3,"Account","",DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR_OCC1)
call VerifyGridCellContent("",4,"Account","",DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KTONR_OCC1)

call VerifyGridCellContent("",2,"Profit Center","",DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR_OCC1)
call VerifyGridCellContent("",3,"Profit Center","",DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_PRCTR_OCC1)


''----------------------Tcode S_ALR_87012357----------------------------
'Enter the Tcode
Call SetTcode(DT_MIGO_0750_OKCD_OCC1) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_MIGO_0750_OKCD_OCC1)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Company code","BR_BUKRS-LOW","",DT_MIGO_1000_COMPANY_CODE,False)
Call SetTextbox("Document Number","BR_BELNR-LOW","",DT_MIGO_1000_DOCUMENT_NUMBER,False)
Call SetTextbox("Fiscal Year","BR_GJAHR-LOW","",DT_MIGO_1000_FISCAL_YEAR,False)
Call SetTextbox("Posting date","BR_BUDAT-LOW","",ConvertDate(Date),False)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButtonIfExist("Execute   \(F8\)",False)
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

Call SendKey("{PGDN}")
Wait(2)
Call SendKey("{PGDN}")
Wait(2)
Call SendKey("{PGDN}")
Wait(2)

Call VerifyifGuiLabelWithIndexExists(DT_MIGO_0120_CHECK_TEXT_OF_RO02,1)
CAll VerifyifGuiLabelWithIndexExists(DT_MIGO_0120_CHECK_TEXT_OF_A2,1)
Call VerifyifGuiLabelWithIndexExists(DT_MIGO_0120_CHECK_TEXT_OF_INPUT_TAX_VAT_9_DEDUCTIBLE,1)

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()


