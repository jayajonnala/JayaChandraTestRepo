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
'.................Test Script Name : Test_Clear AP Accounts - Manual_p3_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 8th April
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Clear AP Accounts - Manual_p3_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Clear AP Accounts - Manual_p3_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
'''----------------------Tcode F-44----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SelectRadioButton("RF05A-XPOS1",DT_F44_2000_DOCUMENT_NUMBER,False)
Call SetTextbox("Account","RF05A-AGKON","",DT_F44_0131_ACCOUNT,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F44_0131_COMPANY_CODE,False)
'Call SetTextbox("Special G/L Ind","RF05A-AGUMS","",DT_F44_0131_SPECIAL_GL_IND,False)
'Call SetTextbox("Clearing Date","BKPF-BUDAT","",Replace(DT_F44_0131_CLEARING_DATE,"/","."),False)

'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
'Call ClickButton("Process Open Items   \(Shift\+F4\)",False)

Wait(1)
'Capture the screenshot
Call TakeScreenShot()

'set filter criteria
Call SetTextbox("From","RF05A-SEL01",0,DT_F44_0731_FROM,False)
Call SetTextbox("From","RF05A-SEL01",1,DT_F44_0731_FROM_OCC1,False)

'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
'Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Select All",False) 
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Deactivate Items",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Field content search",False) 
Call TakeScreenShot()

Call SelectRadioButton("RF05A-XPOS1","Document Number",True)
Call TakeScreenShot()

Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot()

Call SetTextbox("From","RF05A-SEL01","",DT_F44_0731_FROM_OCC2,True)
Call SetTextbox("From","RF05A-SEL01",1,DT_F44_0731_FROM_OCC3,True)'''''''
Call TakeScreenShot()

Call PressEnter() 
Call TakeScreenShot()

Call ClickButton("Select All",False) 
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Activate Items",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Document Overview   \(Shift\+F2\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
wait(1)
'Capture the screenshot
Call TakeScreenShot()
'Validate If doc is generated
Call GetStatusBar("item1","DT_F44_0131_DOCUMENT_NUMBER_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_F44_0131_CHECK_TEXT_OF_STATUSBAR)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

