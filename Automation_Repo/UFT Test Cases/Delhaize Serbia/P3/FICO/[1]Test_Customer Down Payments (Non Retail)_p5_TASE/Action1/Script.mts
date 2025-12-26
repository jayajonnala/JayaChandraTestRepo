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
'.................Test Script Name : Test_Customer Down Payments (Non Retail)_p5_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 20th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Customer Down Payments (Non Retail)_p5_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Customer Down Payments (Non Retail)_p5_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath) 

Call EndDateof445PeriodByDate(DT_TODAY,"DT_ENDING_DATE_PERIOD")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'Increment the parameter
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_REFERENCE",(Cint(DT_INCREMENT_REFERENCE)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

''----------------------Tcode ZFIAR_RS_RFKORD50PDF----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Variant","V-LOW","","ZRS4",True)
Call SetTextbox("Created By","ENAME-LOW","","ZIVANOVAL",True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",True)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Number","RBELNR-LOW","",DT_ZFIAR_RS_RFKORD50PDF_1000_DOCUMENT_NUMBER,False)
Call FocusTextBox("Document Number","RBELNR-LOW",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Wait(5)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Exit   \(Shift\+F3\)",False)
Call ClickButton("Exit   \(Shift\+F3\)",False)
''
'''----------------------Tcode F-39----------------------------
''Enter the Tcode
Call SetTcode(DT_ZFIAR_RS_RFKORD50PDF_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_ZFIAR_RS_RFKORD50PDF_0100_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Date","BKPF-BLDAT","",DT_ZFIAR_RS_RFKORD50PDF_0115_DOCUMENT_DATE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_ZFIAR_RS_RFKORD50PDF_0115_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_ZFIAR_RS_RFKORD50PDF_0115_COMPANY_CODE,False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",DT_ZFIAR_RS_RFKORD50PDF_0115_POSTING_DATE,False)
Call SetTextbox("Period","BKPF-MONAT","",DT_ZFIAR_RS_RFKORD50PDF_0115_PERIOD,False)
Call SetTextbox("Currency","BKPF-WAERS","",DT_ZFIAR_RS_RFKORD50PDF_0115_CURRENCY,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_ZFIAR_RS_RFKORD50PDF_0115_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_ZFIAR_RS_RFKORD50PDF_0115_DOCHEADER_TEXT,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_ZFIAR_RS_RFKORD50PDF_0115_ACCOUNT,False)
Call SetTextbox("Invoice ref\.","BSEG-REBZG","",DT_ZFIAR_RS_RFKORD50PDF_0115_INVOICE_REF,False)
Call SetTextbox("Fiscal year","BSEG-REBZJ","",DT_ZFIAR_RS_RFKORD50PDF_0115_FISCAL_YEAR,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
'Capture the screenshot
Call TakeScreenShot()
wait 3
Call PressEnter() 
'wait 5
'Call PressEnter() 
'wait 5
'Call PressEnter() 
Call GetStatusBar("item1","DT_ZFIAR_RS_RFKORD50PDF_1703_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_ZFIAR_RS_RFKORD50PDF_1703_CHECK_TEXT_OF_STATUSBAR)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC2)
'Capture the screenshot
Call TakeScreenShot()

''selecting cell of row:1 , Column "Document Number"
''SAPGuiSession("transaction:=F-39").SAPGuiWindow("transaction:=F-39").SAPGuiTable("guicomponenttype:=80","name:=SAPMF05ATC_1703").SelectCell 1,"Document Number"
''Call DoubleClick()
'''Capture the screenshot
''Call TakeScreenShot()

Call SelectRowGuiTable("SAPMF05ATC_1703","Document Number",DT_DOCUMENT,False)
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Wait(2)
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
wait 2
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC4)

Call SendKey("{F2}")
Wait(2)
Call SetTextbox("Requested line item","\*BSEG-BUZEI","","2",True)
'''Call ClickButtonIfExist("Continue Processing   \(Shift\+F1\)",True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(2)

Call SetTextbox("Text","BSEG-SGTXT","",DT_ZFIAR_RS_RFKORD50PDF_0301_TEXT,False)
Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()
Call ClickButtonIfExist("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call SendKey("{F2}")
Wait(2)
Call SetTextbox("Requested line item","\*BSEG-BUZEI","","4",True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
' SetTextbox(textboxAttachedText, textboxName, textboxIndex, textboxValue, blnIsItPopup)
Call SetTextbox("Profit Ctrs","COBL-PRCTR",0,DT_PROFIT_CENTER,False)
Call PressEnter()     ' 
Call ClickButtonIfExist("Back   \(F3\)",True)

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Wait(2)

'Capture the screenshot
Call TakeScreenShot()
Call GetStatusBar("item1","DT_ZFIAR_RS_RFKORD50PDF_0115_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetStatusBar("item2","DT_ZFIAR_RS_RFKORD50PDF_0115_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_ZFIAR_RS_RFKORD50PDF_0115_CHECK_TEXT_OF_STATUSBAR)
Call VerifyStatusBarMessageType("S")

CAll SelectMenuBar("Document;Display")
Wait(2)
'Capture the screenshot
Call TakeScreenShot()


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

