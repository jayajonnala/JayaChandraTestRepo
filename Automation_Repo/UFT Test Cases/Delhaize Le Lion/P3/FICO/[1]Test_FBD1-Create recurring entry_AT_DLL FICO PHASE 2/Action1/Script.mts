'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_FBD1-Create recurring entry  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_FBD1-Create recurring entry"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\DLL_P3\Data\TASE_DT_02-04-01-05-03-Create new assortment-BASA.xls"

If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	datatable_row= Parameter("datatable_row")	
End If
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

'Call CloseSessionsSAP()
'SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
'Call Login(DT_SAPUSER,DT_SAPPASSWORD)
'Call PressEnter()

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''''''--------TransactionCode-FBD1----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
GetRowNo =4
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)            
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FBD1_0106_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FBD1_0106_ACCOUNT,False)
Call SetTextbox("Document Header Text","BKPF-BKTXT","",DT_FBD1_0106_DOCUMENT_HEADER_TEXT,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_FBD1_0106_REFERENCE,False)
Call SetTextbox("First run on","BKDF-DBBDT","",ConvertDate(DT_FBD1_0106_FIRST_RUN_ON),False)
Call SetTextbox("Last run on"," BKDF-DBEDT","",ConvertDate(DT_FBD1_0106_LAST_RUN_ON),False)
Call SetTextbox("Interval in months","BKDF-DBMON","",DT_FBD1_0106_INTERVAL_IN_MONTHS,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FBD1_0106_CURRENCYRATE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FBD1_0106_COMPANY_CODE,False)
Call SetTextbox("Document type","BKPF-BLART","",DT_FBD1_0106_DOCUMENT_TYPE,False)
Call TakeScreenShot
Call PressEnter()           
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FBD1_0300_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FBD1_0300_ACCOUNT,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FBD1_0300_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_FBD1_0300_TAX_CODE,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_FBD1_0300_TEXT,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_FBD1_1007_COST_CENTER,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_FBD1_1007_BUSINESS_AREA,False)
Call TakeScreenShot
Call PressEnter() 
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FBD1_0300_AMOUNT_OCC1,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_FBD1_0300_TAX_CODE_OCC1,False)
Call ClickButton("All Acct Assignmts",False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_FBD1_1012_BUSINESS_AREA,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call ClickButton("All Acct Assignmts",False)
Call SetTextbox("Profit Center","COBL-PRCTR","",DT_FBD1_0002_PROFIT_CENTER,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call SetTextbox("Text","BSEG-SGTXT","",DT_FBD1_0300_TEXT_OCC1,False)
Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetStatusBar("item1", "DT_OP_DOC_NO")
GetRowNo =4
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyStatusBar(DT_FBD1_0106_CHECK_TEXT_OF_STATUSBAR)
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
