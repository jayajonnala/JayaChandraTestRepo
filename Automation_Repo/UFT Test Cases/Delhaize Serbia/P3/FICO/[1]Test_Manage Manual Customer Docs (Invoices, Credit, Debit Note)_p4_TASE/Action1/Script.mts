
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Manage Manual Customer Docs (Invoices, Credit, Debit Note)_p4
'.................Author : TCS        :Jaya
'................ Creation Date    : 
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Manage Manual Customer Docs (Invoices, Credit, Debit Note)_p4_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Clear GL Accounts  Manual and Automatic_p1_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''----------------------Tcode FB05----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

Call SelectRadioButton("RF05A-XPOS1",DT_FB05_0122_TRANSFER_POSTING_WITH_CLEARING,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",Replace((DT_FB05_0122_DOCUMENT_DATE),"/","."),False)
Call SetTextbox("Type","BKPF-BLART","",DT_FB05_0122_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB05_0122_COMPANY_CODE,False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",Replace((DT_FB05_0122_POSTING_DATE),"/","."),False)
Call SetTextbox("Period","BKPF-MONAT","",DT_FB05_0122_PERIOD,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FB05_0122_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_FB05_0122_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_FB05_0122_DOCHEADER_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB05_0122_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB05_0122_ACCOUNT,False)

Call TakeScreenShot()
Call PressEnter() 
wait(1)
Call TakeScreenShot()

Call SelectCheckbox("RF05A-XSTBA", 0, DT_FB05_0312_DETERMINE_TAX_BASE, False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_FB05_0312_TAX_CODE,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB05_0312_AMOUNT,False)

Call PressEnter() 
wait(1)
Call TakeScreenShot()

Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB05_0312_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB05_0312_ACCOUNT,False)

Call PressEnter() 
wait(1)
Call TakeScreenShot()

Call SelectCheckbox("RF05A-XSTBA", 0, DT_FB05_0312_DETERMINE_TAX_BASE_OCC1, False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_FB05_0312_TAX_CODE_OCC1,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB05_0312_AMOUNT_OCC1,False)

Call PressEnter() 
wait(1)
Call TakeScreenShot()

Call SelectmenuBar("Document;Simulate")
Call PressEnter() 
wait(1)
Call TakeScreenShot()

Call FocusTextBoxByIndex("PK  BusA Acct                               RSD   Amount        Tax amnt","RF05A-AZEI1", 0, False)
'Call SendKey("{F2}")
Call ClickButton("Choose   \(F2\)",False)
Call TakeScreenShot()
Call VerifyTextBoxContent("Account","SKAT-TXT50","",Lcase(DT_FB05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOBEZ),False)
Call VerifyTextBoxContent("Amount","BSEG-WRBTR","",DT_FB05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET,False)  
Call VerifyTextBoxContent("Tax Code","BSEG-MWSKZ","",DT_FB05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ,False)  

Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot()
Call FocusTextBoxByIndex("PK  BusA Acct                               RSD   Amount        Tax amnt","RF05A-AZEI1", 1, False)
'Call SendKey("{F2}")
Call ClickButton("Choose   \(F2\)",False)
Call TakeScreenShot()
Call VerifyTextBoxContent("Account","SKAT-TXT50","",Lcase(DT_FB05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KOBEZ),False)  
'''Call VerifyTextBoxContent("Amount","BSEG-WRBTR","",DT_FB05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET,False)  
Call VerifyTextBoxContent("Amount","BSEG-WRBTR","",DT_FB05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET,False)  
Call VerifyTextBoxContent("Tax Code","BSEG-MWSKZ","",DT_FB05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MWSKZ,False) 

Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot()

Call PressEnter()
Call ClickButton("Post   \(Ctrl\+S\)",False)
Wait 2
Call TakeScreenShot()
Call PressEnter()
Wait 2
Call TakeScreenShot()
Call GetStatusBar("item1","DT_FB05_0122_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document "&DT_FB05_0122_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" was posted in company code RS01")
Call WriteRunTimeDataToExcelGlobalSheet("DT_FB05_0122_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_FB05_0122_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()


