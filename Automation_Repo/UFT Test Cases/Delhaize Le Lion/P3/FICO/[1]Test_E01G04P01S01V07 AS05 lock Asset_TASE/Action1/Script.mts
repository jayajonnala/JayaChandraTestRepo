'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_E01G04P01S01V07 AS05 lock Asset
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_E01G04P01S01V07 AS05 loc As"
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

''''''--------TransactionCode-AS05----------''''
'
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS05_0100_COMPANY_CODE,False)
Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS05_0100_ASSET,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AS05_0100_SUBNUMBER,False)
Call TakeScreenShot
Call PressEnter()
Call SelectRadioButton("ANLA-XSPEB","Locked to acquis.", False)
Call TakeScreenShot
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call VerifyStatusBar(DT_AS05_0100_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot

''''''--------TransactionCode-F-02----------''''
Call SetTcode(DT_EXPECTEDTRANSACTIONCODE_OCC1)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)


Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_AS05_0100_POSTING_DATE),False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_AS05_0100_REFERENCE,False)
Call SetTextbox("Doc.Header Text","BKPF-BKTXT","",DT_AS05_0100_DOCHEADER_TEXT,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_AS05_0100_CURRENCYRATE,False)
Call SetTextbox("TType","RF05A-NEWBW","",DT_AS05_0100_TTYPE,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_AS05_0100_DOCUMENT_DATE),False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_AS05_0100_ACCOUNT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_AS05_0100_PSTKY,False)
Call SetTextbox("Type","BKPF-BLART","",DT_AS05_0100_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_AS05_0100_COMPANY_CODE_OCC1,False)
Call TakeScreenShot
Call SetTextbox("Period","BKPF-MONAT","",DT_AS05_0100_PERIOD,False)
Call TakeScreenShot
Call TakeScreenShot
Call PressEnter()
cALL VerifyStatusBarMessageType(DT_AS05_0100_MESSAGE_TYPE)
Call VerifyStatusBar(DT_AS05_0100_CHECK_TEXT_OF_STATUSBAR_OCC1)
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
