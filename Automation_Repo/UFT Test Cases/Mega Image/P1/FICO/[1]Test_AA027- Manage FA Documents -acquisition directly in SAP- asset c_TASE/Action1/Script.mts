
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AA027- Manage FA Documents -acquisition directly in SAP- asset c
'.................Author : TCS 
'................ Creation Date :
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


gstrTestCaseName = "Test_AA027- Manage FA Documents -acquisition directly in SAP- asset c"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\FICO\TASE_DT_AA010- Asset Creation - create asset directly in SAP.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)


Call StartDateof445PeriodByDate(DT_TODAY,"DT_START_DATE")
Call EndDateof445PeriodByDate(DT_TODAY,"DT_END_DATE")

'
''.......................Mandatory Initial Call only in First Component in a Test Scenario
'
''''''--------------login----------------'''''
'
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()



'''''--------TransactionCode-F-90----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F90_0100_COMPANY_CODE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_F90_0100_TYPE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F90_0100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F90_0100_ACCOUNT,False)
''Call SetTextbox("Document Date","BKPF-BLDAT","",DT_F90_0100_DOCUMENT_DATE,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F90_0100_DOCUMENT_DATE),False)

Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F90_0100_CURRENCYRATE,False)
Call TakeScreenShot
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F90_0100_REFERENCE,False)
Call PressEnter()  
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F90_0302_AMOUNT,False)
Call SetTextbox("Tax Amount","BSEG-WMWST","",DT_F90_0302_TAX_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F90_0302_TAX_CODE,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F90_0302_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F90_0302_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F90_0302_ACCOUNT,False)
Call SetTextbox("TType","RF05A-NEWBW","",DT_F90_0302_TTYPE,False)
Call SetTextbox("New co.code","RF05A-NEWBK","",DT_F90_0302_NEW_COCODE,False)
Call PressEnter()
Call PressEnter()
Call TakeScreenShot

''Call SetTextbox("Amount","BSEG-WRBTR","",DT_F90_0305_AMOUNT,False)
Call SetTextboxNoLabel("BSEG-WRBTR","",DT_F90_0305_AMOUNT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F90_0305_TEXT,False)
Call ClickButtonIfExist("Display Document Overview   \(Shift\+F2\)",False)
Call TakeScreenShot
Call VerifyTextBoxContent("C","RF05A-AZSAL","",DT_F90_0700_CHECK_TEXT_OF_C,False)
Call PressEnter()
Call TakeScreenShot

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call PressEnter()
Call TakeScreenShot
Call GetStatusBar("item1", "DT_F90_DOC_NUMBER_OUTPUT")
Call VerifyStatusBar("Document "&DT_F90_DOC_NUMBER_OUTPUT&" was posted in company code RO02")
Call TakeScreenShot

Call SetTcode(DT_F90_0750_OKCD)  
Call PressEnter()
Call PressEnter() 

Call VerifyTextBoxContent("Capitalized on","ANLA-AKTIV","",ConvertDate(DT_F90_1142_CHECK_TEXT_OF_CAPITALIZED_ON),False)

Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call TakeScreenShot

Call SelectTab("TABSTRIP100", "Deprec. Areas", False)
Call TakeScreenShot
''Call FindRowNumber("SAPLAISTTC_ANLB","ODep Start",ConvertDate(DT_F90_1190_CHECK_TEXT_OF_TABLECELL_ODEP_START_0),"DT_ROW")
Call VerifyTableCellContent(1, "ODep Start", "SAPLAISTTC_ANLB", ConvertDate(DT_F90_1190_CHECK_TEXT_OF_TABLECELL_ODEP_START_0))
Call VerifyTableCellContent(2, "ODep Start", "SAPLAISTTC_ANLB", ConvertDate(DT_F90_1190_CHECK_TEXT_OF_TABLECELL_ODEP_START_1))
Call VerifyTableCellContent(3, "ODep Start", "SAPLAISTTC_ANLB", ConvertDate(DT_F90_1190_CHECK_TEXT_OF_TABLECELL_ODEP_START_2))
Call VerifyTableCellContent(4, "ODep Start", "SAPLAISTTC_ANLB", ConvertDate(DT_F90_1190_CHECK_TEXT_OF_TABLECELL_ODEP_START_3))
Call VerifyTableCellContent(5, "ODep Start", "SAPLAISTTC_ANLB", ConvertDate(DT_F90_1190_CHECK_TEXT_OF_TABLECELL_ODEP_START_4))

Call LogOff()

Call FinalStatus ()





