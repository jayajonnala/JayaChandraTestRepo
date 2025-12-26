'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Manual Posting of GL Document both ledgers_p2
'.................Author : TCS        :Bitan
'................ Creation Date    : 8th April
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

gstrTestCaseName = "Test_Manual Posting of GL Document both ledgers_p2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Clear AP Accounts - Manual_p4_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''----------------------Login----------------------------
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''''--------TransactionCode-FB01----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB01_0100_COMPANY_CODE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_FB01_0100_TYPE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB01_0100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB01_0100_ACCOUNT,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_FB01_0100_DOCUMENT_DATE),False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_FB01_0100_POSTING_DATE),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FB01_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_FB01_0100_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_FB01_0100_DOCHEADER_TEXT,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB01_0300_AMOUNT,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_FB01_1007_COST_CENTER,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_FB01_1007_BUSINESS_AREA,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB01_0300_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","1",DT_FB01_0300_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB01_0300_AMOUNT,False)
Call TakeScreenShot

Call SelectMenuBar("Document;Simulate")
Call TakeScreenShot

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_FB01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call TakeScreenShot
Call VerifyStatusBar("Document "&DT_FB01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" was posted in company code RS01")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FB01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_FB01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)


Call LogOff()



