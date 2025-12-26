
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_GL0033 Clear GL Accounts (Manual & Automatic) - automatic cleari_TASE
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


gstrTestCaseName = "Test_GL0033 Clear GL Accounts (Manual & Automatic) - automatic cleari_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call StartDateof445PeriodByDate(DT_TODAY,"DT_STARTING_DATE_PERIOD")
''''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
'
''''----------------------Tcode FB01----------------------------
'


Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_FB01_0100_DOCUMENT_DATE),False)
Call SetTextbox("Type","BKPF-BLART","",DT_FB01_0100_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB01_0100_COMPANY_CODE,False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_FB01_0100_POSTING_DATE),False)
Call SetTextbox("Period","BKPF-MONAT","",DT_FB01_0100_PERIOD,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FB01_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_FB01_0100_REFERENCE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB01_0100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB01_0100_ACCOUNT,False)

Call TakeScreenShot
Call PressEnter
Call TakeScreenShot
Call PressEnter

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB01_0300_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_FB01_0300_TAX_CODE,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_FB01_1012_BUSINESS_AREA,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB01_0300_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","1",DT_FB01_0300_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB01_0300_AMOUNT_OCC1,False)
Call ClickButton("COBL_MORE",False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_FB01_1012_BUSINESS_AREA_OCC1,True)
Call TakeScreenShot
'Call PressEnter

Call SetTextbox("Transactn type","COBL-RMVCT","",DT_TRANSACTION_TYPE,False)
Call PressEnter()
Call TakeScreenShot
'

Call SelectMenuBar("Document;Simulate")
Call TakeScreenShot
Call PressEnter()
wait 5
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)

Call GetStatusBar("item1","DT_DOC_NO_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_NO_OUTPUT&" was posted in company code RO02")
Call TakeScreenShot

Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_FB01_0100_DOCUMENT_DATE_OCC1),False)
Call SetTextbox("Type","BKPF-BLART","",DT_FB01_0100_TYPE_OCC1,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB01_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_FB01_0100_POSTING_DATE_OCC1),False)
Call SetTextbox("Period","BKPF-MONAT","",DT_FB01_0100_PERIOD_OCC1,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FB01_0100_CURRENCYRATE_OCC1,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_FB01_0100_REFERENCE_OCC1,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB01_0100_PSTKY_OCC1,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB01_0100_ACCOUNT_OCC1,False)

Call TakeScreenShot
While SAPGuiSession(sessionObject).SapGuiWindow(windowobject).SAPGuiEdit("guicomponenttype:=31","attachedtext:=Reference","name:=BKPF-XBLNR","Index:=0").Exist(5)=True 
Call PressEnter()  
Wend
''Call PressEnter
''Call TakeScreenShot
'''Call PressEnter
'''Call PressEnter
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB01_0300_AMOUNT_OCC2,False)
Call ClickBUtton("All Acct Assignmts",False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_FB01_1012_BUSINESS_AREA_OCC1,True)
Call SetTextbox("Transactn type","COBL-RMVCT","",DT_TRANSACTION_TYPE,False)
Call PressEnter()
Call TakeScreenSHot()
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB01_0300_PSTKY_OCC1,False)
Call SetTextbox("Account","RF05A-NEWKO","1",DT_FB01_0300_ACCOUNT_OCC1,False)
Call TakeScreenShot
Call PressEnter

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB01_0300_AMOUNT_OCC3,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_FB01_0300_TAX_CODE,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_FB01_1012_BUSINESS_AREA_OCC2,False)
Call TakeScreenShot
Call PressEnter

Call SelectMenuBar("Document;Simulate")
Call TakeScreenShot
Call PressEnter()
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)

Call GetStatusBar("item1","DT_DOC_NO_2_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_NO_2_OUTPUT&" was posted in company code RO02")
Call TakeScreenShot
'
'''----------------------Tcode F-03----------------------------

Call SetTcode(DT_FB01_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_FB01_0100_OKCD)

Call SetTextbox("Account","RF05A-AGKON","",DT_FB01_0131_ACCOUNT,False)
Call SetTextbox("Clearing Date","BKPF-BUDAT","",ConvertDate(DT_FB01_0131_CLEARING_DATE),False)
Call SetTextbox("Period","BKPF-MONAT","",DT_FB01_0131_PERIOD,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB01_0131_COMPANY_CODE,False)
Call SetTextbox("Currency","BKPF-WAERS","",DT_FB01_0131_CURRENCY,False)
Call TakeScreenShot
Call SelectradioButton("RF05A-XPOS1","Document Number",False)
Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
Call TakeScreenShot
While SAPGuiSession(sessionObject).SapGuiWindow(windowobject).SAPGuiEdit("guicomponenttype:=32","attachedtext:=Account","name:=RF05A-AGKON","Index:=0").Exist(5)=True 
Call PressEnter()  
Wend
Call SetTextbox("From","RF05A-SEL01",0,DT_DOC_NO_OUTPUT,False)
Call SetTextbox("From","RF05A-SEL01",1,DT_DOC_NO_2_OUTPUT,False)
Call TakeScreenShot
Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
Call PressEnter()
Call TakeScreenShot
Call Selectmenubar("Document;Simulate")
Call TakeScreenShot
Call VerifyStatusBar(Lcase("The difference is too large for clearing"))

Call LogOff()
Call FinalStatus()
