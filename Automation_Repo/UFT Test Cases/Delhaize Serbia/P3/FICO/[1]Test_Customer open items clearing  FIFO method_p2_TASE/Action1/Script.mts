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

'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''
''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
''.................Test Script Name : Test_Customer open items clearing FIFO method_p2_TASE
''.................Author : TCS        :Bitan
''................ Creation Date    : 1st May
''.................Modified By :
''.................Modified Date/Details :
'
''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
gstrTestCaseName = "Test_Customer open items clearing FIFO method_p2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Customer open items clearing  FIFO method_p2_TASE.xls"
'''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
''
'''''Login'''
''DataRowSet=2
'Call StartExecution1(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)
'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
'''----------------------Tcode FB01----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDateFormat(DT_FB01_0100_DOCUMENT_DATE),False)
Call SetTextbox("Type","BKPF-BLART","",DT_FB01_0100_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB01_0100_COMPANY_CODE,False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDateFormat(DT_FB01_0100_POSTING_DATE),False)
Call SetTextbox("Period","BKPF-MONAT","",DT_FB01_0100_PERIOD,False)
Call PressEnter()
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FB01_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_FB01_0100_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_FB01_0100_DOCHEADER_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB01_0100_PSTKY,False)
Call FocusTextBox("Account","RF05A-NEWKO",False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB01_0100_ACCOUNT,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()
'Call PressEnter()
'Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FB01_0100_CURRENCYRATE_OCC1,False)
'Call SetTextbox("Account","RF05A-NEWKO","",DT_FB01_0100_ACCOUNT,False)
'Capture the screenshot
Call TakeScreenShot()

'Call PressEnter()  
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB01_0301_AMOUNT,False)
Call SetTextbox("Bus\. Area","BSEG-GSBER","",DT_FB01_0301_BUS_AREA,False)
Call SetTextbox("Bline Date","BSEG-ZFBDT","",ConvertDateFormat(DT_FB01_0301_BLINE_DATE),False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB01_0301_PSTKY,False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Account","RF05A-NEWKO","",DT_FB01_0301_ACCOUNT,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB01_0300_AMOUNT,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectMenuBar("Document;Simulate")
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyTextBoxContent("Type","BKPF-BLART","",DT_FB01_0750_CHECK_TEXT_OF_TYPE,False)
Call VerifyTextBoxContent("Company Code","RF05A-AZBUK","",DT_FB01_0750_CHECK_TEXT_OF_COMPANY_CODE,False)

Call FocusTextBox("Company Code","RF05A-AZBUK",False)
Call ClickButton("Choose   \(F2\)",False)
Call SetTextbox("Requested line item","\*BSEG-BUZEI","",1,True)
Call ClickButton("Continue   \(Enter\)",True)
Wait(1)
Call VerifyTextBoxContent("G/L Acc","BSEG-HKONT","",DT_FB01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Next Item   \(Shift\+F7\)",False)
Wait(1)
Call VerifyTextBoxContent("G/L Account","BSEG-HKONT","",DT_FB01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Back   \(F3\)",False)
Wait(2)
Call VerifyTextBoxContent("PK  BusA Acct                               RSD   Amount        Tax amnt","RF05A-AZEI1",0,Lcase(DT_VALIDATE_TEXT1),False)
Call VerifyTextBoxContent("PK  BusA Acct                               RSD   Amount        Tax amnt","RF05A-AZEI1",1,Lcase(DT_VALIDATE_TEXT2),False)

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
'Validate If invoice is generated
Call GetStatusBar("item1","DT_DOCUMENT_NUMBER_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_FB01_0100_CHECK_TEXT_OF_STATUSBAR_OCC1)

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

