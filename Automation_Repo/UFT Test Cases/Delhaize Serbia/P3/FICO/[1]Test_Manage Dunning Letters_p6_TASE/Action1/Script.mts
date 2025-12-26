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
'.................Test Script Name : Test_Manage Dunning Letters_p6_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 20th April
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Manage Dunning Letters_p6_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Manage Dunning Letters_p6_TASE.xls"
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
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

''----------------------Tcode FB05----------------------------

'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
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
Call SetTextbox("SGL Ind","RF05A-NEWUM","",DT_FB05_0122_SGL_IND,False)

'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 
wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB05_0304_AMOUNT,False)
Call SetTextbox("Bus\. Area","BSEG-GSBER","",DT_FB05_0304_BUS_AREA,False)
Call SetTextbox("Due on","BSEG-ZFBDT","",Replace((DT_FB05_0304_DUE_ON),"/","."),False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
Call PressEnter() 
wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Company Code","RF05A-AGBUK","",DT_FB05_0710_COMPANY_CODE,False)
Call SetTextbox("Account","RF05A-AGKON","",DT_FB05_0710_ACCOUNT,False)

Call PressEnter() 
wait(1)
'Capture the screenshot
Call TakeScreenShot()

'select and activate respective doc and validation___________________________________________
Call ClickButton("Select All",False) 
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Deactivate Items",False) 
'Capture the screenshot
Call TakeScreenShot()

'Call TakeScreenShot()
Call ClickButton("Field content search",False) 
Call SelectRadioButton("RF05A-XPOS1",DT_FB05_2000_DOCUMENT_NUMBER,True)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True) 
Call SetTextbox("From","RF05A-SEL01","0",DT_FB05_0731_FROM,True)
Call SetTextbox("From","RF05A-SEL01","1",DT_FB05_0731_TO,True)
'Call SetTextbox("To","RF05A-SEL02","",DT_FB05_0731_TO,True)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True) 

'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Select All",False) 
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Activate Items",False) 
'Capture the screenshot
Call TakeScreenShot()

'Call VerifyTextBoxContent("Not assigned","RF05A-DIFFB","",DT_FB05_6102_CHECK_TEXT_OF_NOT_ASSIGNED,False)  'Value:=0,00
Call VerifyTextBoxContent("Not assigned","RF05A-DIFFB","","0,00 ",False) 
'______________________________________________________________________

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()

Call GetTextStatusBar("DT_Verify_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_FB05_0122_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)

'set focus to the faulty item
SAPGuiSession("transaction:=FB05").SAPGuiWindow("transaction:=FB05").SAPGuiEdit("name:=RF05A-AZEI1","attachedtext:=PK  BusA Acct                               RSD   Amount        Tax amnt","index:=2").SetFocus
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Choose   \(F2\)",False) 
Call SetTextbox("Business Area","COBL-GSBER","",DT_FB05_1007_BUSINESS_AREA,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(2)
Call TakeScreenShot()

'Validate If invoice is generated
Call GetStatusBar("item1","DT_FB05_0122_GET_TEXT_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_FB05_0122_CHECK_TEXT_OF_STATUSBAR)

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

