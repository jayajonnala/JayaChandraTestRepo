'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_Customer Clearing_PRE2  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Customer Clearing_PRE2"
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

''--------TransactionCode-FB01----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE) 

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
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB01_0300_AMOUNT,False)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB01_0301_ACCOUNT,False)
Call SetTextbox("Bline Date","BSEG-ZFBDT","",ConvertDate(DT_FB01_0301_BLINE_DATE),False)
Call SetTextboxNoLabel("BSEG-ZTERM","",DT_FB01_0301_PAYT_TERMS,False)
Call SetTextbox("Days/percent","BSEG-ZBD1T","",DT_FB01_0301_DAYSPERCENT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB01_0301_PSTKY,False)
Call TakeScreenShot
Call PressEnter()
Call PressEnter()
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB01_0300_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_FB01_0300_TAX_CODE,False)
Call PressEnter()
Call TakeScreenShot
Call SetTextbox("Business Area","COBL-GSBER","",DT_FB01_1007_BUSINESS_AREA,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_FB01_1007_COST_CENTER,False)
Call ClickButtonIfExist("Display Document Overview   \(Shift\+F2\)",False)
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetStatusBar("item1","DT_OP_DOC_2")
GetRowNo =2
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyStatusBar(DT_FB01_0100_CHECK_TEXT_OF_STATUSBAR)
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
