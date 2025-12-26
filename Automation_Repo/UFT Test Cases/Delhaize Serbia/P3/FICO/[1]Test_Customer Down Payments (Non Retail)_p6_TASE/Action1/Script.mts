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
'.................Test Script Name : Test_Customer Down Payments (Non Retail)_p6_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 21st May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Customer Down Payments (Non Retail)_p6_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Customer Down Payments (Non Retail)_p6_TASE.xls"
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
'Reload
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'''----------------------Tcode F-32----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SelectRadioButton("RF05A-XPOS1",DT_F32_0131_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F32_0131_COMPANY_CODE,False)
Call SetTextbox("Clearing Date","BKPF-BUDAT","",Replace((DT_F32_0131_CLEARING_DATE),"/","."),False)
Call SetTextbox("Account","RF05A-AGKON","",DT_F32_0131_ACCOUNT,False)
Call SetTextbox("Currency","BKPF-WAERS","",DT_F32_0131_CURRENCY,False)
Call SetTextbox("Period","BKPF-MONAT","","",False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("From","RF05A-SEL01","0",DT_F32_0731_FROM,False)
Call SetTextbox("From","RF05A-SEL01","1",DT_F32_0731_FROM_OCC1,False)

Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
VerifyStatusBar(DT_F32_0731_CHECK_TEXT_OF_STATUSBAR)

Call ClickButton("Process Open Items   \(Shift\+F4\)",False)

Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)
Call GetStatusBar("item1","DT_F32_3100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_F32_3100_CHECK_TEXT_OF_STATUSBAR)

'select and activate respective doc and validation___________________________________________
Call ClickButton("Select All",False) 
'Capture the screenshot
Call TakeScreenShot()

''Call ClickButton("Deactivate Items",False) 
Call ClickButton("Activate Items",False) 
'Call SendKey("{F2}")
'Capture the screenshot
Call TakeScreenShot()

'''selecting cell of row:1 , Column "RSD Gross"
''SAPGuiSession("transaction:=F-32").SAPGuiWindow("transaction:=F-32").SAPGuiTable("guicomponenttype:=80","name:=SAPDF05XTC_6102").SelectCell 1,"RSD Gross"
''Call DoubleClick()
''SAPGuiSession("transaction:=F-32").SAPGuiWindow("transaction:=F-32").SAPGuiTable("guicomponenttype:=80","name:=SAPDF05XTC_6102").SelectCell 2,"RSD Gross"
''Call DoubleClick()
'''Capture the screenshot
''Call TakeScreenShot()

''''_________________________________________________________________________________________

Call ClickButton("Post   \(Ctrl\+S\)",False)
''wait 3
''Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
''Call ClickButtonIfExist("Save",True)
wait(2)
Call TakeScreenShot()
'Validate If invoice is generated
Call GetStatusBar("item1","DT_F32_0131_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetStatusBar("item2","DT_F32_0131_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_F32_0131_CHECK_TEXT_OF_STATUSBAR)

Call SelectMenuBar("Document;Display")
wait(2)
Call TakeScreenShot()

Call ClickButton("Display Document Header   \(F5\)",False)
Call TakeScreenShot()

Call ClickButton("Continue/Confirm   \(Enter\)",False)
Call TakeScreenShot()

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

