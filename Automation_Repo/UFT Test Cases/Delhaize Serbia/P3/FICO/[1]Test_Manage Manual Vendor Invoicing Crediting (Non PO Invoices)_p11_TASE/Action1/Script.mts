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
'.................Test Script Name : Test_Manage Manual Vendor Invoicing Crediting (Non PO Invoices)_p11_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 29th April
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Manage Manual Vendor Invoicing Crediting (Non PO Invoices)_p11_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Manage Manual Vendor Invoicing  Crediting (Non PO Invoices)_p11_TASE.xls"
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
'''----------------------Tcode FB02----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FB02_0100_FISCAL_YEAR,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB02_0100_COMPANY_CODE,False)
Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB02_0100_DOCUMENT_NUMBER,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Display Document Header   \(F5\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyTextBoxContent("Document type","BKPF-BLART","",DT_FB02_1710_CHECK_TEXT_OF_DOCUMENT_TYPE,True)
Call VerifyTextBoxContent("Document type","T003T-LTEXT","",Lcase(DT_FB02_1710_CHECK_TEXT_OF_DOCUMENT_TYPE_OCC1),True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Continue/Confirm   \(Enter\)",True)
'Capture the screenshot
Call TakeScreenShot()

Call DoubleClickGuiGridCell("","",1,DT_FB02_0750_GRIDCELL_0_DESCRIPTION,False)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Payt Terms","BSEG-ZTERM","",DT_FB02_0302_PAYT_TERMS,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
Wait(2)
'Capture the screenshot
Call TakeScreenShot()
VerifyStatusBar(DT_FB02_0302_CHECK_TEXT_OF_STATUSBAR)

Call PressEnter() 
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Wait(2)
VerifyStatusBar(DT_FB02_0100_CHECK_TEXT_OF_STATUSBAR)

Call PressEnter() 
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call DoubleClickGuiGridCell("","",1,DT_FB02_0750_GRIDCELL_0_DESCRIPTION_OCC1,False)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyTextBoxContent("Payt Terms","BSEG-ZTERM","",DT_FB02_0302_CHECK_TEXT_OF_PAYT_TERMS,False)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyTextBoxContent("Days/percent","BSEG-ZBD1T","",DT_FB02_0302_CHECK_TEXT_OF_DAYSPERCENT,False)

Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

