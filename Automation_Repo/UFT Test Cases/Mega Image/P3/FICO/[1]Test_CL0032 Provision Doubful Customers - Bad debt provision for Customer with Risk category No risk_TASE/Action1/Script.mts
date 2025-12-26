'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_CL0032 Provision Doubful Customers - Bad debt provision for Customer with Risk category No risk
'.................Author : TCS      
'................ Creation Date   
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


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


gstrTestCaseName = "Test_CL0032 "
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\FICO\TASE_DT_CL0032.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''----------------------Login----------------------------

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
'''''--------TransactionCode---------''''

Call SetTcode("fbl5n")     
Call PressEnter()     
Call TakeScreenShot


Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_CUSTOMER,False)
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_COMP_CO,False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)", Falser)
Call TakeScreenShot

'''''--------TransactionCode-FD32---------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call SelectCheckbox("RF02L-D0210", 0, DT_FD32_0100_STATUS, False)
Call SelectCheckbox("RF02L-D0220", 0, DT_FD32_0100_PAYMENT_HISTORY, False)
Call SetTextbox("Customer","RF02L-KUNNR","",DT_FD32_0100_CUSTOMER,False)
Call SetTextbox("Credit control area","RF02L-KKBER","",DT_FD32_0100_CREDIT_CONTROL_AREA,False)
Call TakeScreenShot
Call PressEnter() 
Call Wait(5)
Call FocusTextBox("Risk category", "KNKK-CTLPC", False)
Call SendKey("{F4}")
Call SetFocusGuiLabel("No risk", 81, 72, True)
Call SendKey("{F2}")


'''''--------TransactionCode-ZFIAR_F103---------''''

Call SetTcode(DT_FD32_0100_OKCD)     
Call PressEnter() 

Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_FD32_1000_CUSTOMER_ACCOUNT,False)
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_FD32_1000_COMPANY_CODE,False)
Call SetTextbox("Key date","STICHTAG","",ConvertDate(DT_FD32_1000_DOCUMENT_DATE),False)
Call SetTextbox("Provision method","RSMET","",DT_FD32_1000_PROVISION_METHOD,False)
Call SetTextbox("Spec.G/L indic.for dbtfl. rec.","B-UMSKZ","",DT_FD32_1000_SPECGL_INDICFOR_DBTFL_REC,False)
Call SetTextbox("Posting date","B-BUDAT","",ConvertDate(DT_FD32_1000_POSTING_DATE),False)
Call SetTextbox("Document date","B-BLDAT","",ConvertDate(DT_FD32_1000_DOCUMENT_DATE),False)
Call SetTextbox("Posting period","B-BUPEM","",DT_FD32_1000_POSTING_PERIOD,False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)", Falser)
Call TakeScreenShot

'''''--------TransactionCode-ZFIAR_F104---------''''

Call SetTcode(DT_FD32_0100_OKCD_OCC1)     
Call PressEnter() 
Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_FD32_1000_CUSTOMER_ACCOUNT,False)
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_FD32_1000_COMPANY_CODE,False)
Call SetTextbox("Open items at key date","DD_STIDA","",ConvertDate(DT_FD32_1000_KEY_DATE_OCC1),False)
Call SetTextbox("Key date","STICHTAG","",ConvertDate(DT_FD32_1000_KEY_DATE_OCC1),False)
Call TakeScreenShot
Call FocusTextBox("Valuation Area", "BWBER", False)
Call SendKey("{F4}")

Call ClickButtonIfExist("Copy   \(Enter\)", True)
Call SetTextbox("Posting date","B_BUDAT","",ConvertDate(DT_FD32_1000_KEY_DATE_OCC1),False)
Call SetTextbox("Document date","B_BLDAT","",ConvertDate(DT_FD32_1000_KEY_DATE_OCC1),False)
Call SetTextbox("Posting period","B_BUPEM","",DT_FD32_1000_POSTING_PERIOD_OCC1,False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)", False)
Call TakeScreenShot


Call LogOff()

Call FinalStatus ()

