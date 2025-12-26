
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_GL0028_Clear_GL_Accounts_Manual_Automatic_standard_manual_TASE
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


gstrTestCaseName = "Test_GL0028_Clear_GL_Accounts_Manual_Automatic_standard_manual_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call StartDateof445PeriodByDate(DT_TODAY,"DT_STARTING_DATE_PERIOD")
'
''''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

''''''''--------TransactionCode-F-02----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'''Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_F02_100_POSTING_DATE),False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",DT_F02_100_POSTING_DATE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F02_100_COMPANY_CODE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_F02_100_TYPE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_100_ACCOUNT,False)
''Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F02_100_DOCUMENT_DATE),False)
Call SetTextbox("Document Date","BKPF-BLDAT","",DT_F02_100_DOCUMENT_DATE,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F02_100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F02_100_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F02_100_DOCHEADER_TEXT,False)

Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Text","BSEG-SGTXT","",DT_F02_300_TEXT,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_300_AMOUNT,False)
Call SetTextbox("Tax code","BSEG-MWSKZ","",DT_F02_300_TAX_CODE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_300_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_300_ACCOUNT,False)

Call TakeScreenShot

Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_300_AMOUNT_OCC2,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F02_300_TEXT_OCC2,False)

Call TakeScreenShot
Call ClickButton("All Acct Assignmts",False)
Call SetTextbox("Transactn type","COBL-RMVCT","",DT_TRANSACTION_TYPE,False)
Call PressEnter()
Call TakeScreenShot

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call PressEnter()
Call GetStatusBar("item1","DT_DOC_NO1_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_NO1_OUTPUT&" was posted in company code RO02")

Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_NO1_OUTPUT",DT_DOC_NO1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'

Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_F02_100_POSTING_DATE_OCC2),False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F02_100_COMPANY_CODE_OCC2,False)
Call SetTextbox("Type","BKPF-BLART","",DT_F02_100_TYPE_OCC2,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_100_PSTKY_OCC2,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_100_ACCOUNT_OCC2,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F02_100_DOCUMENT_DATE_OCC2),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F02_100_CURRENCYRATE_OCC2,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F02_100_REFERENCE_OCC2,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F02_100_DOCHEADER_TEXT_OCC2,False)
Call SetTextbox("TType","RF05A-NEWBW","",DT_TRANSACTION_TYPE,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
'Call PressEnter()

Call SetTextbox("Text","BSEG-SGTXT","",DT_F02_300_TEXT_OCC3,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_300_AMOUNT_OCC3,False)

Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_300_PSTKY_OCC2,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_300_ACCOUNT_OCC2,False)

Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_300_AMOUNT_OCC4,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F02_300_TEXT_OCC4,False)
Call SetTextbox("Tax code","BSEG-MWSKZ","",DT_F02_300_TAX_CODE_OCC2,False)

Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call PressEnter()
Call GetStatusBar("item1","DT_DOC_NO2_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_NO2_OUTPUT&" was posted in company code RO02")

Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_NO2_OUTPUT",DT_DOC_NO2)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

''''''''''--------TransactionCode-F-03----------''''
'
Call SetTcode(DT_F02_100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)


Call SetTextbox("Account","RF05A-AGKON","",DT_F02_131_ACCOUNT,False)
Call SetTextbox("Clearing Date","BKPF-BUDAT","",ConvertDate(DT_CLEARING_DATE),False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F02_131_COMPANY_CODE,False)
Call SetTextbox("Currency","BKPF-WAERS","",DT_F02_131_CURRENCY,False)
'''Call  SelectRadioButton("RF05A-XPOS1","Document Number",False)
Call TakeScreenShot
Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
Call TakeScreenShot
Call PressEnter()

Call ClickButton("Select All",False)
Call TakeScreenShot
Call ClickButton("Deactivate Items",False)
Call TakeScreenShot
Call ClickButton("Field content search",False)
Call TakeScreenShot

Call SelectRadioButton("RF05A-XPOS1", "Document Number", True)

Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SetTextbox("From","RF05A-SEL01",0,DT_F02_0731_FROM_OCC3,False)
Call SetTextbox("From","RF05A-SEL01",1,DT_F02_0731_FROM_OCC4,False)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot

Call ClickButton("Select All",False)
Call TakeScreenShot
Call ClickButton("Activate Items",False)
Call TakeScreenShot

Call ClickButton("Document Overview   \(Shift\+F2\)",False)
Call TakeScreenShot

Call SelectMenuBar("Document;Simulate")
Call ClickButton("Choose   \(F2\)",False)
Call SetTextbox("Requested line item","\*BSEG-BUZEI",0,DT_LINE_ITEM,False)
Call ClickButton("Continue   \(Enter\)",False)
Call TakeScreenShot
Call ClickButton("All Acct Assignmts",False)
Call SetTextbox("Transactn type","COBL-RMVCT","",DT_TRANSACTION_TYPE,False)
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Next Item   \(Shift\+F7\)",False)
'Call ClickButton("All Acct Assignmts",False)
''updating as property value changed
Call ClickButton("All acct assignments",False)
Call SetTextbox("Transactn type","COBL-RMVCT","",DT_TRANSACTION_TYPE,False)
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)",False)
'Call PressEnter()   
'Call TakeScreenShot
'Call PressEnter()   
While SAPGuiSession(sessionObject).SapGuiWindow(windowobject).SAPGuiEdit("guicomponenttype:=32","attachedtext:=Document Date","name:=BKPF-BLDAT","Index:=0").Exist(5)=True 
Call PressEnter()  
Wend
Call TakeScreenShot
Call GetStatusBar("item1","DT_DOC_NO3_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_NO3_OUTPUT&" was posted in company code RO02")

Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_NO3_OUTPUT",DT_DOC_NO3)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'''''''--------TransactionCode-FAGLL03----------''''

Call SetTcode(DT_F02_131_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)

Call SetTextbox("G/L account","SD_SAKNR-LOW",0,DT_F02_1000_GL_ACCOUNT,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call ClickLabel("DocumentNo",1, False)

Call ClickButtonifexist("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call TakeScreenShot
Call ClickButtonifexist("Find",True)
Call TakeScreenShot
''''''Call SetTextbox("Find","GD_SEARCHSTR",0,"Document Number",Ture)
''''''Call TakeScreenShot
''''''Call ClickButtonifexist("Continue   \(Enter\)",True)
''''''Call TakeScreenShot
''''''Call ClickButtonifexist("Show sel\. fields \(CTRL\+F3\)",True)
''''''Call TakeScreenShot
''''''Call ClickButtonifexist("Copy   \(Enter\)",True)
''''''Call TakeScreenShot
Call SetTextbox("Document Number","%%DYN001-LOW",0,DT_F02_1105_CLEARING_DOCUMENT,True)
Call TakeScreenShot
Call ClickButtonifexist("Execute   \(Enter\)",True)
Call TakeScreenShot

Call VerifyifGuiLabelExists(LCase(DT_F02_1000_CHECK_TEXT_OF_STATUSBAR))

Call ClickButton("Back   \(F3\)",FAlse)

Call SelectRadioButton("X_AISEL", "All Items", False)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call ClickLabel("DocumentNo",1, False)

Call ClickButtonifexist("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call TakeScreenShot
''Call ClickButtonifexist("Find",True)
''Call TakeScreenShot
''Call SetTextbox("Find","GD_SEARCHSTR",0,"Document Number",True)
''Call TakeScreenShot
''Call ClickButtonifexist("Continue   \(Enter\)",True)
''Call TakeScreenShot
''Call ClickButtonifexist("Show sel\. fields \(CTRL\+F3\)",True)
''Call TakeScreenShot
''Call ClickButtonifexist("Copy   \(Enter\)",True)
''Call TakeScreenShot
Call clickbutton("Multiple selection",True)
Call TakeScreenShot
Call SetTableDataNoRef("SAPLALDBSINGLE", "Single value", 1, DT_DOC_NO1, True)
Call SetTableDataNoRef("SAPLALDBSINGLE", "Single value", 2, DT_DOC_NO2, True)
Call SetTableDataNoRef("SAPLALDBSINGLE", "Single value", 3, DT_DOC_NO3, True)
Call TakeScreenShot
Call clickbutton("Copy   \(F8\)",True)
Call TakeScreenShot
Call TakeScreenShot
Call ClickButtonifexist("Execute   \(Enter\)",True)
Call TakeScreenShot

Call VerifyifGuiLabelExists(DT_DOC_NO1)
Call VerifyifGuiLabelExists(DT_DOC_NO2)
Call VerifyifGuiLabelExists_ByIndex(DT_DOC_NO3,0)
''''Call VerifyifGuiLabelExistsByRelativeid(DT_DOC_NO3,"wnd\[0\]/usr/lbl\[9,10\]")
''''Call VerifyifGuiLabelExistsByRelativeid(DT_F02_120_CHECK_TEXT_OF_SA," wnd\[0\]/usr/lbl\[25,8\]")
''''Call VerifyifGuiLabelExistsByRelativeid(DT_F02_120_CHECK_TEXT_OF_SA_OCC2, "wnd\[0\]/usr/lbl\[25,9\]")


Call VerifyifGuiLabelExists_ByIndex(DT_F02_120_CHECK_TEXT_OF_SA,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F02_120_CHECK_TEXT_OF_SA_OCC2,1)
''Call VerifyifGuiLabelExists_ByIndex(DT_F02_120_CHECK_TEXT_OF_CL)
Call VerifyifGuiLabelExists_ByIndex(DT_F02_120_CHECK_TEXT_OF_CL,0)

''Call VerifyifGuiLabelExists(DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC4)
Call VerifyifGuiLabelExists(DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC5)
''Call VerifyifGuiLabelExists(DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC6)
Call VerifyifGuiLabelExists_ByIndex(DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC6,0)
''''Call VerifyifGuiLabelExistsByRelativeid(DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC8, "wnd\[0\]/usr/lbl\[40,9\]")
''''Call VerifyifGuiLabelExistsByRelativeid(DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC8, "wnd\[0\]/usr/lbl\[40,10\]")


Call VerifyifGuiLabelExists_ByIndex(DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC8,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC8,1)
''Call VerifyifGuiLabelExists(DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC7)
Call VerifyifGuiLabelExists_ByIndex(DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC7,0)

Call VerifyifGuiLabelExists(DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC9)
Call VerifyifGuiLabelExists(DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC10)
'''Call VerifyifGuiLabelExists(DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC11)
'''Call VerifyifGuiLabelExists_ByIndex(DT_F02_120_CHECK_TEXT_OF_NO_NAME_OCC11,0)
'

Call ClickLabel(DT_DOC_NO3, "", False)
wait 5
Call ClickButton("Continue   \(Enter\)",True)
Call ClickButton("Display Document   \(Ctrl\+Shift\+F7\)",False)
Wait 5
Call TakeScreenSHot()
Call ClickButton("Call Up Document Overview   \(F9\)",False)
Wait 5
Call TakeScreenSHot()
Call DoubleClickGuiGridCell("", "", 1, "Account", False)

Call VerifyTextBoxContent("G/L Account", "SKAT-TXT50", "",Lcase( DT_F02_300_CHECK_TEXT_OF_GL_ACCOUNT_OCC2), False)
Call VerifyTextBoxContent("G/L Account", "BSEG-HKONT", "", DT_F02_300_CHECK_TEXT_OF_GL_ACCOUNT, False)
Call VerifyTextBoxContent("Amount", "BSEG-WRBTR", "", DT_F02_300_CHECK_TEXT_OF_AMOUNT, False)
Call ClickButton("Back   \(F3\)",FAlse)
wait 5
Call ClickLabel(DT_DOC_NO3, "", False)
Call PressEnter()
Call ClickButton("Display Document   \(Ctrl\+Shift\+F7\)",False)
Wait 5
Call TakeScreenSHot()

Call VerifyTextBoxContent("G/L Account", "SKAT-TXT50", "", Lcase(DT_F02_300_CHECK_TEXT_OF_GL_ACCOUNT_OCC31), False)
Call VerifyTextBoxContent("G/L Account", "BSEG-HKONT", "", DT_F02_300_CHECK_TEXT_OF_GL_ACCOUNT_OCC3, False)
'Call VerifyTextBoxContent("Amount", "BSEG-WRBTR", "", DT_F02_300_CHECK_TEXT_OF_AMOUNT_OCC2, False)

Call Logoff'
Call FinalStatus()
