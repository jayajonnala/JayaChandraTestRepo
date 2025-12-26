'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_PRE_01_E10G06P02S01V03 Clear GL Accounts (Automatic with currenc 
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_PRE_01_ClearGL(withCur)"
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

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''--------TransactionCode-FB01----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE) 

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",DT_INCREMENT+1)  
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call SetTextbox("Period","BKPF-MONAT","",DT_FB01_0100_PERIOD,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_FB01_0100_DOCUMENT_DATE),False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_FB01_0100_POSTING_DATE),False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB01_0100_COMPANY_CODE,False)
Call SetTextbox("Doc.Header Text","BKPF-BKTXT","",DT_FB01_0100_DOCHEADER_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB01_0100_PSTKY,False)
Call SetTextbox("Type","BKPF-BLART","",DT_FB01_0100_TYPE,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB01_0100_ACCOUNT,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_FB01_0100_REFERENCE,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FB01_0100_CURRENCYRATE,False)
Call TakeScreenShot
Call PressEnter()

'''If VerifyStatusBarMessageType("W") Then
'''Call TakeScreenShot
'''Call PressEnter()
'''Call TakeScreenShot
'''End if


Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB01_0300_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_FB01_0300_TAX_CODE,False)
Call ClickButton("All Acct Assignmts",False)
Call TakeScreenShot
Call SetTextbox("Profit Center","COBL-PRCTR","",DT_FB01_1005_PROFIT_CENTER,True)
Call ClickButton("Continue   \(Enter\)",True)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB01_0300_ACCOUNT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB01_0300_PSTKY,False)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB01_0300_AMOUNT_OCC1,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_FB01_0300_TAX_CODE_OCC1,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_FB01_1007_COST_CENTER,False)
Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetTextStatusBar("DT_OP_FB01_0100_OP_CHECK_TEXT_OF_STATUSBAR")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)

'''''''--------TransactionCode-/FB03----------''''


Call SetTcode(DT_FB01_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)

'Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB01_0100_DOCUMENT_NUMBER,False)
Call SetTextboxNoLabel("RF05L-BELNR","",DT_FB01_0100_DOCUMENT_NUMBER,False)
'Call SetTextbox("Doc. Number","RF05L-BELNR","",DT_FB01_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB01_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FB01_0100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()
Call VerifyTextBoxContent("Reference","BKPF-XBLNR","", DT_FB01_0750_CHECK_TEXT_OF_REFERENCE, False)
Call ClickButton("Display Document Header   \(F5\)",False)
Call VerifyTextBoxContent("Document type","BKPF-BLART","", DT_FB01_1710_CHECK_TEXT_OF_DOCUMENT_TYPE, True)
Call GetTextboxValue("BKPF-KURSF","","DT_OP_FB01_1710_OP_CHECK_TEXT_OF_EXCHANGE_RATE", True)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call ClickButtonIfExist("Continue/Confirm   \(Enter\)",True)
Call DoubleClickGuiGridCell("", 0, 1, "Account", False)
Call TakeScreenShot
Call VerifyTextBoxContent("Amt.in loc.cur.","BSEG-DMBTR","",DT_FB01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMBTR, False)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
