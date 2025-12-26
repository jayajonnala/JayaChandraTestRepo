'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_PRE_03_E10G07P02S01V01 Client Listing for BE10  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_PRE_03_E10G07P02S01V01 Listing for BE10"
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
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''--------TransactionCode-F-04----------''''


Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F04_0122_COMPANY_CODE,False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_F04_0122_POSTING_DATE),False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F04_0122_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F04_0122_ACCOUNT,False)
Call SetTextbox("Type","BKPF-BLART","",DT_F04_0122_TYPE,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F04_0122_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F04_0122_REFERENCE,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F04_0122_DOCUMENT_DATE),False)
Call SetTextbox("Doc.Header Text","BKPF-BKTXT","",DT_F04_0122_DOCHEADER_TEXT,False)
Call TakeScreenShot
Call PressEnter()     
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F04_0300_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F04_0300_ACCOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F04_0300_TAX_CODE,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F04_0300_AMOUNT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F04_0300_TEXT,False)
Call TakeScreenShot
Call PressEnter() 
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F04_0300_PSTKY_OCC1,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F04_0300_ACCOUNT_OCC1,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F04_0300_TAX_CODE_OCC1,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F04_0300_AMOUNT_OCC1,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F04_0300_TEXT_OCC1,False)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("Text","BSEG-SGTXT","",DT_F04_0301_TEXT,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F04_0301_AMOUNT,False)
Call TakeScreenShot
Call ClickButton("Display Document Overview   \(Shift\+F2\)", False)
Call TakeScreenShot
Call ClickButton("Enter/Display Tax Items   \(F5\)", False)
Call TakeScreenShot
Call ClickButton("Calculate   \(F7\)", False)
Call SetTextbox("Amount","BSET-FWSTE",1,DT_F04_0300_AMOUNT_OCC2,False)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)", False)
Call ClickButton("Post   \(Ctrl\+S\)", False)
'Call PressEnter()
'Commenting PressEnter as it NR
Call TakeScreenShot
Call GetStatusBar("item1", "DT_OP_F04_0122_CHECK_MESSAGEPARAMETER_OF_STATUSBAR")
GetRowNo = 4
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyStatusBar(DT_F04_0122_CHECK_TEXT_OF_STATUSBAR)
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
