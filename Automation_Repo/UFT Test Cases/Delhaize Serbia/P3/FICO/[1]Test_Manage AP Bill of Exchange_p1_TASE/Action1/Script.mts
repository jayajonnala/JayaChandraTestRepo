'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Manage AP Bill of Exchange_p1_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 10th March
'.................Modified By :
'.................Modified Date/Details :
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

gstrTestCaseName = "Test_Manage AP Bill of Exchange_p1_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Manage AP Bill of Exchange_p1_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
'''----------------------Tcode FB05----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SelectRadioButton("RF05A-XPOS1",DT_FB05_0122_TRANSFER_POSTING_WITH_CLEARING,False)
Call SetTextbox("Period","BKPF-MONAT","",DT_FB05_0122_PERIOD,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB05_0122_COMPANY_CODE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB05_0122_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB05_0122_ACCOUNT,False)
Call SetTextbox("Type","BKPF-BLART","",DT_FB05_0122_TYPE,False)
Call SetTextbox("SGL Ind","RF05A-NEWUM","",DT_FB05_0122_SGL_IND,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",Replace((DT_FB05_0122_DOCUMENT_DATE),"/","."),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FB05_0122_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_FB05_0122_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_FB05_0122_DOCHEADER_TEXT,False)

'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 
wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB05_2320_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB05_2320_ACCOUNT,False)
Call SetTextbox("Due on","BSEG-ZFBDT","",Replace((DT_FB05_2320_DUE_ON),"/","."),False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_FB05_2320_TEXT,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB05_2320_AMOUNT,False)

'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB05_0300_AMOUNT,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectMenuBar("Document;Simulate")
'Capture the screenshot
Call TakeScreenShot()

''if script fails with bellow line where the value is hard coaded,,, any of the next three lines to be "un commented" or have to change the hardcoded value
Call VerifyTextBoxContent("PK  BusA Acct                               RSD   Amount        Tax amnt","RF05A-AZEI1",0,"001 39E      0010030933 STANČIĆ DOO               100.000,00-",False)
'Call VerifyTextBoxContent("PK  BusA Acct                               RSD   Amount        Tax amnt","RF05A-AZEI1",0,DT_FB05_0700_CHECK_TEXT_OF_PK_BUSA_ACCT_text_1,False)
'Call VerifyTextBoxContent("PK  BusA Acct                               RSD   Amount        Tax amnt","RF05A-AZEI1",0,UCase(DT_FB05_0700_CHECK_TEXT_OF_PK_BUSA_ACCT_text_1),False)
'Call VerifyTextBoxContent("PK  BusA Acct                               RSD   Amount        Tax amnt","RF05A-AZEI1",0,LCase(DT_FB05_0700_CHECK_TEXT_OF_PK_BUSA_ACCT_text_1),False)

''text_2
Call VerifyTextBoxContent("PK  BusA Acct                               RSD   Amount        Tax amnt","RF05A-AZEI1",1,LCASE(DT_FB05_0700_CHECK_TEXT_OF_PK_BUSA_ACCT_text_2),False)

'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call PressEnter()
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
'Call PressEnter()
'Validate If invoice is generated
Call GetStatusBar("item1","DT_FB05_0122_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_FB05_0122_CHECK_TEXT_OF_STATUSBAR)

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

