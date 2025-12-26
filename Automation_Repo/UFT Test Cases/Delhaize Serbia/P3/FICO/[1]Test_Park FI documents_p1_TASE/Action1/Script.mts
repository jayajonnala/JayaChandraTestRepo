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
''Reload DataSheet to updates and calculations
''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Park FI documents_p1_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 11th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Park FI documents_p1_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Park FI documents_p1_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
'
'''----------------------Tcode FV50----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()
	
'this step is not in the log but it is a mandate in screen
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FV50_1000_COMPANY_CODE,True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Date","ACGL_HEAD-BLDAT","",Replace((DT_FV50_1010_DOCUMENT_DATE),"/","."),False)
Call SetTextbox("Reference","ACGL_HEAD-XBLNR","",DT_FV50_1010_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","ACGL_HEAD-BKTXT","",DT_FV50_1010_DOCHEADER_TEXT,False)
Call SetTextbox("Document type","ACGL_HEAD-BLART","",DT_FV50_1010_DOCUMENT_TYPE,False)
Call SetTextbox("Currency","ACGL_HEAD-WAERS","",DT_FV50_1010_CURRENCY,False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTableDataNoRef("SAPLFSKBTABLE","G/L acct",1,DT_FV50_0100_TABLECELL_GL_ACCT_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","G/L acct",2,DT_FV50_0100_TABLECELL_GL_ACCT_1,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","D/C",1,DT_FV50_0100_TABLECELL_DC_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","D/C",2,DT_FV50_0100_TABLECELL_DC_1,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Amount in doc.curr.",1,DT_FV50_0100_TABLECELL_AMOUNT_IN_DOCCURR_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Amount in doc.curr.",2,DT_FV50_0100_TABLECELL_AMOUNT_IN_DOCCURR_1,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Business area",1,DT_FV50_0100_TABLECELL_BUSINESS_AREA_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Business area",2,DT_FV50_0100_TABLECELL_BUSINESS_AREA_1,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Cost center",1,DT_FV50_0100_TABLECELL_COST_CENTER_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Cost center",2,DT_FV50_0100_TABLECELL_COST_CENTER_1,False)

Call PressEnter() 
wait(1)
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

