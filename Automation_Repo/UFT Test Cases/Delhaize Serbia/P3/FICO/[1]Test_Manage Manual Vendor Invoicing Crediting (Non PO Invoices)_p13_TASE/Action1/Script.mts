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
'.................Test Script Name : Test_Manage Manual Vendor Invoicing Crediting (Non PO Invoices)_p13_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 29th April
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Manage Manual Vendor Invoicing Crediting (Non PO Invoices)_p13_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Manage Manual Vendor Invoicing  Crediting (Non PO Invoices)_p13_TASE.xls"
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
'''----------------------Tcode FBL3N----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_FBL3N_1000_COMPANY_CODE,False)
Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_FBL3N_1000_GL_ACCOUNT,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Wait(2)
Call TakeScreenShot()
CAll ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(10)
Call VerifyStatusBarMessageType("S")
 
Call SelectMenuBar("Edit;Find")
Call SetTextbox("Find","RSYSF-STRING","",DT_FBL3N_1105_DOCUMENT_NUMBER,True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Find   \(Enter\)",True)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()
Call ClickLabel(DT_FBL3N_1105_DOCUMENT_NUMBER,"", True)
Wait(2)

Call ClickLabel(DT_FBL3N_1105_DOCUMENT_NUMBER,"", False)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Call Up Document Overview   \(F9\)",False)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyGridCellContent("",1,"Posting Key","",DT_PK)
Call VerifyGridCellContent("",1,"Account","",DT_GL_ACCOUNT)

Call ClickButton("Back   \(F3\)",False)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()

'''----------------------Tcode FBL3N----------------------------
'Enter the Tcode
Call SetTcode(DT_FBL3N_1000_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_FBL3N_1000_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Created By","ENAME-LOW","","",True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",True)
'Capture the screenshot
Call TakeScreenShot()
Call SelectRowGuiGridbyRowNo("Variant Catalog for Program RFUMSV00","",DT_FBL3N_0600_GRIDCELL_4_VARIANT_NAME,False)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Choose   \(F2\)",True)

Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Posting date","BR_BUDAT-LOW","",Replace((DT_FBL3N_1000_POSTING_DATE),"/","."),False)
Call SetTextbox("to","BR_BUDAT-HIGH","",Replace((DT_FBL3N_1000_TO),"/","."),False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Program run date","PAR_LAUD","",Replace((DT_Run_Date),"/","."),False)
Call SetTextbox("Identification","PAR_LAUI","",DT_Identification,False)
Call FocusTextBox("Identification","PAR_LAUI",False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Fiscal Year","BR_GJAHR-LOW","",DT_Fiscal_year,False)
Call FocusTextBox("to","BR_BUDAT-HIGH",False)
Wait(2)
Call PressEnter() 
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Wait(60)
'Capture the screenshot
Call TakeScreenShot()

Call SelectMenuBar("Edit;Find")
Call SetTextbox("Find","RSYSF-STRING","","Balance per company code",True)
Call ClickButton("Find   \(Enter\)",True)
Wait(2)
Call ClickLabel("Balance per company code","", True)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

