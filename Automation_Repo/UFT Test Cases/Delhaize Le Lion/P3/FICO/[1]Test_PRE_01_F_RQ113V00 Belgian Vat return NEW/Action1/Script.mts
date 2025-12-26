'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_PRE_01_F_RQ113V00 Belgian Vat return NEW  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_PRE_01_F_RQ113V00 Belgian Vat return NEW"
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

'''''--------TransactionCode-FB60----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

'Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB60_1000_COMPANY_CODE,True)
'Call TakeScreenShot
'Call PressEnter()
Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_FB60_1000_COMPANY_CODE)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
'Call SetComboByKey("RF05A-BUSCS",DT_FB60_1100_TRANSACTN)
Call SetTextboxNoLabel("INVFO-ACCNT","",DT_FB60_0010_VENDOR,False)
'This function SetTextboxNoLabel is used for vendor to supplier change.
Call SetTextbox("Invoice date","INVFO-BLDAT","",ConvertDate(DT_FB60_0010_INVOICE_DATE),False)
'Call SetComboByKey("INVFO-BLART",DT_FB60_0010_DOCUMENT_TYPE)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",DT_INCREMENT+1)
'GetRowNo = datatable_row
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call SetTextbox("Reference","INVFO-XBLNR","",DT_FB60_0010_REFERENCE,False)   
Call SetTextbox("Posting Date","INVFO-BUDAT","",ConvertDate(DT_FB60_0010_POSTING_DATE),False)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB60_0010_AMOUNT,False)
Call TakeScreenShot
'Call SetTextbox("Amount","INVFO-WAERS","",DT_FB60_0010_AMOUNT_OCC1,False)
'Call SetTextbox("Tax Amount","INVFO-WMWST","",DT_FB60_0010_TAX_AMOUNT,False)
Call SetTableData("SAPLFSKBTABLE", "G/L acct", 1, "", "", DT_FB60_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", 1, "", "", DT_FB60_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "Tax code", 1, "", "", DT_FB60_0100_TABLECELL_TAX_CODE_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", 1, "", "", DT_FB60_0100_TABLECELL_COST_CENTER_0, False)
Call TakeScreenShot
Call PressEnter()
Call ClickButtonIfExist("Continue   \(Enter\)", True)
Call SelectTab("TS", "Tax", False)
Call ClickButtonIfExist("Continue   \(Enter\)", True)
Call SetTableData("SAPLTAX1STEUER_CTRL", "Base Amount", 1, "", "",DT_FB60_0010_AMOUNT, False)
Call PressEnter()
Call ClickButtonIfExist("Continue   \(Enter\)", True)
Call SelectTab("TS", "Basic data", False)
Call ClickButtonIfExist("Continue   \(Enter\)", True)
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)", False)
Call PressEnter()
Call GetStatusBar("item1", "DT_OP_DOCUMENT_CREATED")
'GetRowNo = datatable_row
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call VerifyStatusBar(DT_DOCUMENT_TO_VERIFY)

'''''''--------TransactionCode-/FB03----------''''
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call SetTcode(DT_FB60_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB60_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB60_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FB60_0100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "Posting Key", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "Posting Key", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 3, "Posting Key", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("", 4, "Posting Key", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BSCHL)
Call VerifyGridCellContent("", 1, "Account", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "Account", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 3, "Account", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
Call VerifyGridCellContent("", 4, "Account", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KTONR)
Call VerifyGridCellContent("", 1, "Tax Code", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ)
Call VerifyGridCellContent("", 2, "Tax Code", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MWSKZ)
Call VerifyGridCellContent("", 3, "Tax Code", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_MWSKZ)
Call VerifyGridCellContent("", 4, "Tax Code", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_MWSKZ)
Call VerifyGridCellContent("", 1, "Amount", 0,DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("", 2, "Amount", 0,DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("", 3, "Amount", 0,DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AZBET)
Call VerifyGridCellContent("", 4, "Amount", 0,DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_AZBET)
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
