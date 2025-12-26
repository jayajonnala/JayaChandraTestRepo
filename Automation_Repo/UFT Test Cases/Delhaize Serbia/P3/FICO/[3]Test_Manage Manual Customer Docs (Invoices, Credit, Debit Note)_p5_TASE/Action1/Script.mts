
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Manage Manual Customer Docs (Invoices, Credit, Debit Note)_p5
'.................Author : TCS        :Jaya
'................ Creation Date    : 
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Manage Manual Customer Docs (Invoices, Credit, Debit Note)_p5_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Clear GL Accounts  Manual and Automatic_p1_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''----------------------Tcode ZFIAR_RS_RFKORD50PDF----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_F37_0100_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call FocusTextBox("Document Number","RBELNR-LOW",False)
Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call SetTextbox("Variant","V-LOW","",DT_VARIANT,True)

'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",True)

Call SetTextbox("Document Number","RBELNR-LOW","",DT_ZFIAR_RS_RFKORD50PDF_1000_DOCUMENT_NUMBER,False)
Call FocusTextBox("Document Number","RBELNR-LOW",False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()

Call ClickButton("Execute   \(F8\)",False)
Wait(5)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Exit   \(Shift\+F3\)",False)
Call TakeScreenShot()
Call ClickButton("Exit   \(Shift\+F3\)",False)
Call TakeScreenShot()

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()
