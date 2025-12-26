'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_PRE_07_E10G07P02S01V01 Rep on Gen Tax Docs from BELUX for BE10 
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_PRE_07_E10G07P02S01V01 BELUX for BE10"
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

'Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''--------TransactionCode-FB41----------''''


Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextbox("Period","BKPF-MONAT","",DT_FB41_0100_PERIOD,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB41_0100_COMPANY_CODE,False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_FB41_0100_POSTING_DATE),False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB41_0100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB41_0100_ACCOUNT,False)
Call SetTextbox("Type","BKPF-BLART","",DT_FB41_0100_TYPE,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FB41_0100_CURRENCYRATE,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_FB41_0100_DOCUMENT_DATE),False)
Call TakeScreenShot
Call PressEnter()
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
GetRowNo = 4
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
If VerifyStatusBarMessageType("W") = True then 
	Call  PressEnter() 
End if
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB41_0300_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB41_0300_ACCOUNT,False)
Call SetTextbox("Value date","BSEG-VALUT","",ConvertDate(DT_FB41_0300_VALUE_DATE),False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB41_0300_AMOUNT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_FB41_0300_TEXT,False)
Call TakeScreenShot
Call PressEnter() 
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_FB41_1007_COST_CENTER,False)
Call SetTextbox("Value date","BSEG-VALUT","",ConvertDate(DT_FB41_0300_VALUE_DATE_OCC1),False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_FB41_0300_TAX_CODE,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB41_0300_AMOUNT_OCC1,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_FB41_0300_TEXT_OCC1,False)
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)", False)
Call TakeScreenShot
Call GetStatusBar("item1", "DT_OP_DOC1")
GetRowNo = 4
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyStatusBar(DT_FB41_0100_CHECK_TEXT_OF_STATUSBAR)

'''''''--------TransactionCode-/FB03----------''''


Call SetTcode(DT_FB41_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)

Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB41_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB41_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FB41_0100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "Tax Code", 0, "")
Call VerifyGridCellContent("", 2, "Tax Code", 0, DT_FB41_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MWSKZ)
Call VerifyGridCellContent("", 2, "Profit Center", 0, DT_FB41_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KOSTL)

'''''--------TransactionCode-FB41----------''''


Call SetTcode(DT_FB41_0750_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call SetTextbox("Period","BKPF-MONAT","",DT_FB41_0100_PERIOD_OCC1,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB41_0100_COMPANY_CODE_OCC2,False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_FB41_0100_POSTING_DATE_OCC1),False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB41_0100_PSTKY_OCC1,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB41_0100_ACCOUNT_OCC1,False)
Call SetTextbox("Type","BKPF-BLART","",DT_FB41_0100_TYPE_OCC1,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FB41_0100_CURRENCYRATE_OCC1,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_FB41_0100_DOCUMENT_DATE_OCC1),False)
Call TakeScreenShot
Call PressEnter()
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
GetRowNo = 4
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
If VerifyStatusBarMessageType("W") = True then 
	Call  PressEnter() 
End if
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB41_0300_PSTKY_OCC1,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB41_0300_ACCOUNT_OCC1,False)
Call SetTextbox("Value date","BSEG-VALUT","",ConvertDate(DT_FB41_0300_VALUE_DATE_OCC2),False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB41_0300_AMOUNT_OCC2,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_FB41_0300_TEXT_OCC2,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_FB41_0300_TAX_CODE_OCC1,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_FB41_1007_COST_CENTER_OCC1,False)
Call TakeScreenShot
Call PressEnter() 
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB41_0300_ACCOUNT_OCC2,False)
Call SetTextbox("Value date","BSEG-VALUT","",ConvertDate(DT_FB41_0300_VALUE_DATE_OCC3),False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB41_0300_PSTKY_OCC2,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB41_0300_AMOUNT_OCC3,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_FB41_0300_TEXT_OCC3,False)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB41_0300_AMOUNT_OCC4,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_FB41_0300_TEXT_OCC4,False)
Call SetTextbox("Value date","BSEG-VALUT","",ConvertDate(DT_FB41_0300_VALUE_DATE_OCC4),False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_FB41_0300_TAX_CODE_OCC2,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB41_0300_PSTKY_OCC3,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB41_0300_ACCOUNT_OCC3,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_FB41_1007_COST_CENTER_OCC2,False)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB41_0300_AMOUNT_OCC5,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_FB41_0300_TEXT_OCC5,False)
Call SetTextbox("Value date","BSEG-VALUT","",ConvertDate(DT_FB41_0300_VALUE_DATE_OCC5),False)
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)", False)
Call TakeScreenShot
Call GetStatusBar("item1", "DT_OP_DOC_2")
GetRowNo = 4
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyStatusBar(DT_FB41_0100_CHECK_TEXT_OF_STATUSBAR_OCC1)

'''''''--------TransactionCode-/FB03----------''''


Call SetTcode(DT_FB41_0100_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)

Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB41_0100_DOCUMENT_NUMBER_OCC1,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB41_0100_COMPANY_CODE_OCC3,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FB41_0100_FISCAL_YEAR_OCC1,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "Tax Code", 0, DT_FB41_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ_OCC1)
Call VerifyGridCellContent("", 2, "Tax Code", 0, "")
Call VerifyGridCellContent("", 3, "Tax Code", 0, DT_FB41_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_MWSKZ)
Call VerifyGridCellContent("", 4, "Tax Code", 0, "")
Call VerifyGridCellContent("", 1, "Profit Center", 0, DT_FB41_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOSTL)
Call VerifyGridCellContent("", 2, "Profit Center", 0, "")
Call VerifyGridCellContent("", 3, "Profit Center", 0, DT_FB41_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KOSTL)
Call VerifyGridCellContent("", 4, "Profit Center", 0, "")
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
