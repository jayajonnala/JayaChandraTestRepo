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


'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Manage Manual Vendor Invoicing Crediting (Non PO Invoices)_p14_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 30th April
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Manage Manual Vendor Invoicing Crediting (Non PO Invoices)_p14_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Manage Manual Vendor Invoicing  Crediting (Non PO Invoices)_p14_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
'''----------------------Tcode XK03----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Company Code","RF02K-BUKRS","",DT_XK03_0101_COMPANY_CODE,False)
Call SetTextbox("Vendor","RF02K-LIFNR","",DT_XK03_0101_VENDOR,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Select All   \(F7\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Enter",False)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Next screen   \(F8\)",False)
Wait(2)
Call ClickButton("Next screen   \(F8\)",False)
Wait(2)
Call ClickButton("Next screen   \(F8\)",False)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyTextBoxContent("Recon\. account","LFB1-AKONT","",DT_XK03_0210_CHECK_TEXT_OF_RECON_ACCOUNT,False)
Call ClickButton("Next screen   \(F8\)",False)
Wait(2)
Call ClickButton("Next screen   \(F8\)",False)
Wait(2)
Call ClickButton("Next screen   \(F8\)",False)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("SPOP-OPTION1",True)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

