'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'................."Test_Vendor block bankruptcy" 
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Vend block bankrupt"
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

Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''--------TransactionCode-XK02----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextboxNoLabel("RF02K-LIFNR","",DT_XK02_0101_VENDOR,False)
''This function SetTextboxNoLabel is used for vendor to supplier change.
Call SetTextbox("Company Code","RF02K-BUKRS","",DT_XK02_0101_COMPANY_CODE,False)
Call SetTextbox("Purchasing Organization","RF02K-EKORG","",DT_XK02_0101_PURCH_ORGANIZATION,False)
Call TakeScreenShot
Call ClickButton("Select All   \(F7\)",False)
Call PressEnter()
Call SetTextbox("Comments","ADDR1_DATA-REMARK","",DT_XK02_0301_COMMENTS,False)
Call SelectMenuBar("Extras;Block Data")
Call TakeScreenShot
Call SelectCheckbox("LFA1-SPERM",1,DT_XK02_0510_ALL_PURCHASING_ORGANIZATIONS, False)
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)", False)
Call TakeScreenShot
Call PressEnter()

'''''--------TransactionCode-XK03----------''''

Call SetTcode(DT_XK02_0101_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextboxNoLabel("RF02K-LIFNR","",DT_XK02_0101_VENDOR_OCC1,False)
''This function SetTextboxNoLabel is used for vendor to supplier change.
Call SetTextbox("Company Code","RF02K-BUKRS","",DT_XK02_0101_COMPANY_CODE_OCC1,False)
Call SetTextbox("Purchasing Organization","RF02K-EKORG","",DT_XK02_0101_PURCH_ORGANIZATION_OCC1,False)
Call TakeScreenShot
Call ClickButton("Deselect All   \(F8\)",False)
Call SelectCheckbox("RF02K-D0215",1,DT_XK02_0510_CHECK_SELECTED_OF_ALL_PURCHASING_ORGANIZATIONS_OCC1, False)
Call PressEnter()
Call TakeScreenShot
Call PressEnter()    
Call ClickButton("Yes",True)
Call TakeScreenShot
Call SetTextboxNoLabel("RF02K-LIFNR","",DT_XK02_0101_VENDOR_OCC2,False)
''This function SetTextboxNoLabel is used for vendor to supplier change.
Call SetTextbox("Company Code","RF02K-BUKRS","",DT_XK02_0101_COMPANY_CODE_OCC2,False)
Call SetTextbox("Purchasing Organization","RF02K-EKORG","",DT_XK02_0101_PURCH_ORGANIZATION_OCC2,False)
Call TakeScreenShot
Call ClickButton("Select All   \(F7\)",False)
Call PressEnter()
Call VerifyTextBoxContent("Comments","ADDR1_DATA-REMARK","",DT_XK02_0301_CHECK_TEXT_OF_COMMENTS,False)
Call SelectMenuBar("Extras;Block Data")
Call TakeScreenShot

'''''--------TransactionCode-XK02----------''''

Call SetTcode(DT_XK02_0510_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextboxNoLabel("RF02K-LIFNR","",DT_XK02_0101_VENDOR_OCC2,False)
''This function SetTextboxNoLabel is used for vendor to supplier change.
Call SetTextbox("Company Code","RF02K-BUKRS","",DT_XK02_0101_COMPANY_CODE_OCC2,False)
Call SetTextbox("Purchasing Organization","RF02K-EKORG","",DT_XK02_0101_PURCH_ORGANIZATION_OCC2,False)
Call TakeScreenShot
Call ClickButton("Select All   \(F7\)",False)
Call PressEnter()
Call SetTextbox("Comments","ADDR1_DATA-REMARK","",DT_XK02_0301_COMMENTS_OCC1,False)
Call SelectMenuBar("Extras;Block Data")
Call TakeScreenShot
Call SelectCheckbox("LFA1-SPERM",1,DT_XK02_0510_ALL_PURCHASING_ORGANIZATIONS_OCC1, False)
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)", False)
Call TakeScreenShot
Call PressEnter()
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
