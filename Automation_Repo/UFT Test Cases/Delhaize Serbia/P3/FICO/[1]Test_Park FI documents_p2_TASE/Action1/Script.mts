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

''Reload DataSheet to updates and calculations
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Park FI documents_p2_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 11th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Park FI documents_p2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Park FI documents_p2_TASE.xls"
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
'''----------------------Tcode FV70----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FV70_1000_COMPANY_CODE,True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Continue   \(Enter\)",True)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Customer","INVFO-ACCNT","",DT_FV70_0510_CUSTOMER,False)
Call SetTextbox("Invoice date","INVFO-BLDAT","",Replace((DT_FV70_0510_INVOICE_DATE),"/","."),False)
Call SetTextbox("Reference","INVFO-XBLNR","",DT_FV70_0510_REFERENCE,False)
Call SetTextbox("Posting Date","INVFO-BUDAT","",Replace((DT_FV70_0510_POSTING_DATE),"/","."),False)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FV70_0510_AMOUNT,False)
Call SetTextbox("Amount","INVFO-WAERS","",DT_FV70_0510_AMOUNT_OCC1,False)

'Select Calculate Tax field as true
Call SelectCheckbox("INVFO-XMWST",0,DT_FV70_0510_CALCULATE_TAX,False)
Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

'set table data
Call SetTableDataNoRef("SAPLFSKBTABLE","G/L acct",1,DT_FV70_0100_TABLECELL_GL_ACCT_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Amount in doc.curr.",1,DT_FV70_0100_TABLECELL_AMOUNT_IN_DOCCURR_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Tax code",1,DT_FV70_0100_TABLECELL_TAX_CODE_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Business area",1,DT_FV70_0100_TABLECELL_BUSINESS_AREA_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Cost center",1,DT_FV70_0100_TABLECELL_COST_CENTER_0,False)
Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

'Click on Post Buton
Call ClickButton("Save parked document   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
'Validate If document is posted and get the status bar nummber
Call GetStatusBar("item1","DT_DOCUMENT_CREATED_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_DOCUMENT_TO_VERIFY)
'Capture the screenshot
Call TakeScreenShot()


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

