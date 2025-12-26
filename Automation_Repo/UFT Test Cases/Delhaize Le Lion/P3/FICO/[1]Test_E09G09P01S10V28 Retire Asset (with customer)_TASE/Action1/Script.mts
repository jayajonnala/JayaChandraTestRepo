'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_E09G09P01S10V28 Retire Asset (with customer)
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_E09G09P01S10V28 Retire Asset (with)"
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

'''''--------------login----------------'''''

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''''--------TransactionCode-F-92----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)

Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_F92_0100_POSTING_DATE),False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F92_0100_REFERENCE,False)
Call SetTextbox("Doc.Header Text","BKPF-BKTXT","",DT_F92_0100_DOCHEADER_TEXT,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F92_0100_CURRENCYRATE,False)
Call SetTextbox("TType","RF05A-NEWBW","",DT_F02_0100_TTYPE,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F92_0100_DOCUMENT_DATE),False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F92_0100_PSTKY,False)
Call SetTextbox("Type","BKPF-BLART","",DT_F92_0100_TYPE,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F92_0100_ACCOUNT,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F92_0100_COMPANY_CODE,False)
Call SetTextbox("Period","BKPF-MONAT","",DT_F92_0100_PERIOD,False)
Call TakeScreenShot
Call PressEnter()

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F92_0301_AMOUNT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F92_0301_TEXT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F92_0301_TAX_CODE,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F92_0301_ACCOUNT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F92_0301_PSTKY,False)
Call SetTextbox("Bus. Area","BSEG-GSBER","",DT_F92_0301_BUS_AREA,False)
Call SelectCheckbox("BKPF-XMWST",0,DT_F92_0301_CALCULATE_TAX, False)
Call TakeScreenShot
Call PressEnter()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F92_0300_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F92_0300_TAX_CODE,False)
Call SelectCheckbox("RF05A-XAABG",0,DT_F92_0300_ASST_RETIREMENT, False)
Call TakeScreenShot
Call SetTextbox("Business Area","COBL-GSBER","",DT_F92_1007_BUSINESS_AREA,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_F92_1007_COST_CENTER,False)
Call TakeScreenShot
Call PressEnter()

Call SelectCheckbox("ANBZ-XVABG",0,DT_F92_0210_COMPLETE_RETIREMENT, True) 
Call SetTextbox("Asset","RA01B-ANLN1","",DT_F92_0210_ASSET,True)
Call SetTextbox("Transaction Type","ANBZ-BWASL","",DT_F92_0210_TRANSACTION_TYPE,True)
Call SetTextbox("Asset Value Date","ANBZ-BZDAT","",ConvertDate(DT_F92_0210_ASSET_VALUE_DATE),True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)", True)
Call TakeScreenShot

Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)",False)  
Call TakeScreenShot
Call GetStatusBar("item1","DT_OP_DOC_NO")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call VerifyStatusBar(DT_F92_0100_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot

''''''--------TransactionCode-/FB03----------''''


Call SetTcode(DT_F92_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)


Call SetTextbox("Document Number","RF05L-BELNR","",DT_F92_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_F92_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_F92_0100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()

Call VerifyGridCellContent("", 1, "Account",0,DT_F92_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "Account",0,DT_F92_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 3, "Account",0,DT_F92_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
Call VerifyGridCellContent("", 4, "Account",0,DT_F92_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KTONR)
Call VerifyGridCellContent("", 5, "Account",0,DT_F92_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_KTONR)
Call VerifyGridCellContent("", 6, "Account",0,DT_F92_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_KTONR)

Call VerifyGridCellContent("", 1, "Amount",0,DT_F92_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("", 2, "Amount",0,DT_F92_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("", 3, "Amount",0,DT_F92_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AZBET)
Call VerifyGridCellContent("", 4, "Amount",0,DT_F92_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_AZBET)
Call VerifyGridCellContent("", 5, "Amount",0,DT_F92_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_AZBET)
Call VerifyGridCellContent("", 6, "Amount",0,DT_F92_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_AZBET)

Call VerifyGridCellContent("", 1, "Profit Center",0,"")
Call VerifyGridCellContent("", 2, "Profit Center",0,DT_F92_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)
Call VerifyGridCellContent("", 3, "Profit Center",0,"")
Call VerifyGridCellContent("", 4, "Profit Center",0,DT_F92_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_PRCTR)
Call VerifyGridCellContent("", 5, "Profit Center",0,DT_F92_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_PRCTR)
Call VerifyGridCellContent("", 6, "Profit Center",0,DT_F92_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_PRCTR)
Call TakeScreenShot

'''''--------TransactionCode-AW01----------''''

Call SetTcode(DT_F92_0750_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call VerifyGridCellContent("Planned values IFRS APC, depreciation",1,"Change","","")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",1,"Year-end","","")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",4,"Change","","")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",4,"Year-end","","")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",11,"Change","","")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",11,"Year-end","","")
Call TakeScreenShot

Call SetTextbox("Asset","ANLA-ANLN1","",DT_F92_0201_ASSET,False)
Call SetTextbox("Asset","ANLA-ANLN2","",DT_F92_0201_ASSET_OCC1,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_F92_0201_COMPANY_CODE,False)
Call SetTextbox("Fiscal year","EDIT_JAHRE","",DT_F92_0202_FISCAL_YEAR,False)

Call SelectNodeGuiTree(0,"Depreciation Areas;D1 Non-Leading (Local);06 Local GAAP APC, depreciation")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",1,"Change","","")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",1,"Year-end","","")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",4,"Change","","")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",4,"Year-end","","")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",11,"Change","","")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",11,"Year-end","","")
Call TakeScreenShot

Call DoubleClickGuiGridCell("Transactions", "", 2, "Amount posted", False)

'''''''--------TransactionCode-/FB03----------''''
     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)
'
'''''''--------TransactionCode-/faglb03----------''''
Call SetTcode(DT_F92_0750_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)

Call SetTextbox("Account Number","RACCT-LOW","",DT_F92_1000_ACCOUNT_NUMBER,False)
Call SetTextbox("Company Code","RBUKRS-LOW","",DT_F92_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RYEAR","",DT_F92_1000_FISCAL_YEAR,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call DoubleClickGuiGridCell("", "",Month(Date)+1, "Period", False)
Call VerifyGridCellContent("", 1, "Amount in local currency","",DT_F92_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 1, "Profit Center","",DT_F92_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("", 1, "Segment","",DT_F92_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SEGMENT)
'
'''''''--------TransactionCode-/fd10n----------''''
Call SetTcode(DT_F92_0111_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC5)

Call SetTextbox("Customer","SO_KUNNR-LOW","",DT_F92_1000_CUSTOMER,False)
Call SetTextbox("Company Code","SO_BUKRS-LOW","",DT_F92_1000_COMPANY_CODE_OCC1,False)
Call SetTextbox("Fiscal year","GP_GJAHR","",DT_F92_1000_FISCAL_YEAR_OCC1,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)

Call DoubleClickGuiGridCell("","",Month(Date)+1,"Period",False)

'Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
'Call ClickButton("Find",True)
'Call SetTextbox("Find","GD_SEARCHSTR","","DOCUMENT NUMBER",True)
'Call ClickButton("Continue   \(Enter\)",True)
'Call ClickButton("Show sel. fields \(CTRL\+F3\)",True)
'Call ClickButton("Copy   \(Enter\)",True)
'Call SetTextbox("Document Number","%%DYN001-LOW","",DT_F92_0100_DOCUMENT_NUMBER,True)
'Call TakeScreenShot
'Call ClickButton("Execute   \(Enter\)",True)
'Call TakeScreenShot

Call ClickButton("Set filter   \(Ctrl\+Shift\+F2\)",False)
Call SelectRowGuiGridbyRowNo("Column Set", 0,3, True)
Call ClickButton("Add Filter Criterion \(F7\)",True)
Call TakeScreenShot
Call ClickButton("Define Filter Values",True)
Call TakeScreenShot
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_F92_0100_DOCUMENT_NUMBER,True)
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot
'
Call VerifyGridCellContent("",1,"Amount in local currency","",DT_F92_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB_OCC1)
Call VerifyGridCellContent("",1,"Cleared/open items symbol","",DT_F92_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ICO_AUGP)

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

