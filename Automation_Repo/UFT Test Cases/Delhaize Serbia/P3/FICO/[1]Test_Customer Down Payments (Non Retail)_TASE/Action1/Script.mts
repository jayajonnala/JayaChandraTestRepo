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
'''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Customer Down Payments (Non Retail)_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 19th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Customer Down Payments (Non Retail)_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Customer Down Payments (Non Retail)_TASE.xls"
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
'Increment the parameter
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_REFERENCE",(Cint(DT_INCREMENT_REFERENCE)+1))
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_DOCHEHEADER",(Cint(DT_INCREMENT_DOCHEHEADER)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'''----------------------Tcode F-37----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

''Call SetTextbox("Document Date","BKPF-BLDAT","",DT_F37_0113_DOCUMENT_DATE,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",Replace(DT_F37_0113_DOCUMENT_DATE,"/","."),False)
Call SetTextbox("Type","BKPF-BLART","",DT_F37_0113_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F37_0113_COMPANY_CODE,False)
''Call SetTextbox("Posting Date","BKPF-BUDAT","",DT_F37_0113_POSTING_DATE,False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",Replace(DT_F37_0113_POSTING_DATE,"/","."),False)
Call SetTextbox("Period","BKPF-MONAT","",DT_F37_0113_PERIOD,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F37_0113_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F37_0113_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F37_0113_DOCHEADER_TEXT,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F37_0113_ACCOUNT,False)
Call SetTextbox("Trg\.sp\.G/L ind\.","RF05A-ZUMSK","",DT_F37_0113_TRGSPGL_IND,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 
Wait 2
Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("New Item   \(F8\)",False)
Call SelectCheckbox("RF05A-XMWST","1",DT_F37_0304_CALCULATE_TAX,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F37_0304_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F37_0304_TAX_CODE,False)
Call SetTextbox("Tax Amount","BSEG-WMWST","",DT_F37_0304_TAX_AMOUNT,False)
Call SetTextbox("Bus\. Area","BSEG-GSBER","",DT_F37_0304_BUS_AREA,False)
Call SetTextbox("Due On","BSEG-ZFBDT","",Replace(DT_F37_0304_DUE_ON,"/","."),False)

Call TakeScreenShot()

Call ClickButtonIfExist("Display Long Texts for Item",False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("1st line","EENO_DYNP-ZEILE","",DT_F37_1001_1ST_LINE,True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButtonIfExist("Copy text   \(F5\)",True)

'Capture the screenshot
Call TakeScreenShot()
Call SelectCheckbox("RF05A-XMWST","1",DT_F37_0304_CALCULATE_TAX_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Display Document Overview   \(Shift\+F2\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyTextBoxContent("PK  BusA Acct                               RSD   Amount        Tax amnt","RF05A-AZEI1","",DT_F37_0700_CHECK_TEXT_OF_DATA,False)

Call ClickButton("Post   \(Ctrl\+S\)",False)
Wait(3)
'Validate If document is posted and get the status bar nummber
Call GetStatusBar("item1","DT_F37_0113_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetStatusBar("item2","DT_F37_0113_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_F37_0113_CHECK_TEXT_OF_STATUSBAR)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)

Call ClickButton("Exit   \(Shift\+F3\)",False)
Call ClickButtonIfExist("Yes",True)
'
'''----------------------Tcode ZFIAR_RS_RFKORD50PDF----------------------------
'Enter the Tcode
Call SetTcode(DT_F37_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_F37_0100_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call FocusTextBox("Document Number","RBELNR-LOW",False)
Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Variant","V-LOW","","ZRS3",True)
Call SetTextbox("Created By","ENAME-LOW","","",True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",True)

Call SetTextbox("Document Number","RBELNR-LOW","",DT_F37_1000_DOCUMENT_NUMBER,False)
Call FocusTextBox("Document Number","RBELNR-LOW",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Wait(5)
'Capture the screenshot
Call TakeScreenShot()


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

