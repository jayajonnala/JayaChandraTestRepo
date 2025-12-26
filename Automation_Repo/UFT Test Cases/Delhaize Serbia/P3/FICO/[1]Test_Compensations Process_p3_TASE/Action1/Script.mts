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
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Compensations Process_p3_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 24th March
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Compensations Process_p3_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Compensations Process_p3_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_REFERENCE",(Cint(DT_INCREMENT_REFERENCE)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)


'''----------------------Tcode F-30----------------------------
'

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

Call SetTextbox("Document Date","BKPF-BLDAT","",Replace((DT_F30_0122_DOCUMENT_DATE),"/","."),False)
Call SetTextbox("Type","BKPF-BLART","",DT_F30_0122_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F30_0122_COMPANY_CODE,False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",Replace((DT_F30_0122_POSTING_DATE),"/","."),False)
Call SetTextbox("Period","BKPF-MONAT","",DT_F30_0122_PERIOD,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F30_0122_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F30_0122_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F30_0122_DOCHEADER_TEXT,False)

Call TakeScreenShot()

Call ClickButton("Choose open items   \(F6\)",False)
Call TakeScreenShot()

Call SelectRadioButton("RF05A-XPOS1",DT_F30_0710_DOCUMENT_NUMBER,False)
Call TakeScreenShot()

Call SetTextbox("Company Code","RF05A-AGBUK","",DT_F30_0710_COMPANY_CODE,False)
Call SetTextbox("Account","RF05A-AGKON","",DT_F30_0710_ACCOUNT,False)
Call SetTextbox("Account Type","RF05A-AGKOA","",DT_F30_0710_ACCOUNT_TYPE,False)
Call TakeScreenShot()

Call PressEnter() 
Call TakeScreenShot()

Call SetTextbox("From","RF05A-SEL01",0,DT_F30_0731_FROM,False)
Call TakeScreenShot()

Call PressEnter() 
Call TakeScreenShot()

CAll VerifyStatusBarMessageType(DT_EXPECTEDVALUE)

Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
Call TakeScreenShot()

CAll VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)

Call ClickButton("Select All",False)
Call ClickButton("Deactivate Items",False)
Call TakeScreenShot()

Call SelectCellGuiTable("SAPDF05XTC_6102","RSD Gross","Document Number",DT_F30_0731_FROM,False)
Call DoubleClick()
Call TakeScreenShot()

Call ClickButton("Cancel   \(F12\)",False)
Call TakeScreenShot()

Call ClickButton("Choose open items   \(F6\)",False)
Call TakeScreenShot()

Call SelectRadioButton("RF05A-XPOS1",DT_F30_0710_DOCUMENT_NUMBER_OCC1,False)
Call TakeScreenShot()

Call SetTextbox("Company Code","RF05A-AGBUK","",DT_F30_0710_COMPANY_CODE_OCC1,False)
Call SetTextbox("Account","RF05A-AGKON","",DT_F30_0710_ACCOUNT_OCC1,False)
Call SetTextbox("Account Type","RF05A-AGKOA","",DT_F30_0710_ACCOUNT_TYPE_OCC1,False)
Call TakeScreenShot()

Call PressEnter() 
Call TakeScreenShot()

Call SetTextbox("From","RF05A-SEL01",0,DT_F30_0731_FROM_OCC1,False)
Call TakeScreenShot()

Call PressEnter() 
Call TakeScreenShot()

CAll VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC2)

Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
Call TakeScreenShot()

CAll VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC3)

Call SelectCellGuiTable("SAPDF05XTC_6102","RSD Gross","Document Number",DT_F30_0731_FROM_OCC1,False)
Call DoubleClick()
Call DoubleClick()
Call TakeScreenShot()

Call ClickButton("Select All",False)
Call ClickButton("Activate Items",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
Call GetStatusBar("item1","DT_F30_0122_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetStatusBar("item2","DT_F30_0122_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call VerifyStatusBarMessageType("S")
Call VerifyStatusBar(DT_F30_0122_CHECK_TEXT_OF_STATUSBAR)

Call LogOff()
Call FinalStatus ()


