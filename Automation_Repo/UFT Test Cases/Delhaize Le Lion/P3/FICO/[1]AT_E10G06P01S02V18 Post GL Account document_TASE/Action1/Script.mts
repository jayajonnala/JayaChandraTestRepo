'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................AT_E10G06P01S02V18 Post GL Account document 
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "AT_E10G06P01S02V18 Post GL Acc_TASE"
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

''--------TransactionCode-FB50L----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)


Call PressEnter()
Call ClickButtonIfExist("Switch Company Code   \(F7\)",False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB50L_1000_COMPANY_CODE,True)
Call TakeScreenShot
Call PressEnter()
Call ClickButtonIfExist("Switch Company Code   \(F7\)",False)
Call SetTextbox("Company Code","BKPF-BUKRS","","LU54",True)
Call TakeScreenShot
Call PressEnter()
Call ClickButtonIfExist("Switch Company Code   \(F7\)",False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB50L_1000_COMPANY_CODE,True)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("Document type","ACGL_HEAD-BLART","",DT_FB50L_1010_DOCUMENT_TYPE,False)
Call SetTextbox("Posting Date","ACGL_HEAD-BUDAT","",ConvertDate(DT_FB50L_1010_POSTING_DATE),False)
Call SetTextbox("Ledger Grp"," ACGL_HEAD-LDGRP","",DT_FB50L_1010_LEDGER_GRP,False)
Call SetTextbox("Doc.Header Text"," ACGL_HEAD-BKTXT","",DT_FB50L_1010_DOCHEADER_TEXT,False)
Call SetTextbox("Currency","ACGL_HEAD-WAERS","",DT_FB50L_1010_CURRENCY,False)
Call SetTextbox("Reference","ACGL_HEAD-XBLNR","",DT_FB50L_1010_REFERENCE,False)
Call SetTextbox("Document Date","ACGL_HEAD-BLDAT","",ConvertDate(DT_FB50L_1010_DOCUMENT_DATE),False)
Call TakeScreenShot
Call SetTableData("SAPLFSKBTABLE", "G/L acct", 1, "", "",DT_FB50L_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "G/L acct", 2, "", "",DT_FB50L_0100_TABLECELL_GL_ACCT_1, False)
Call SetTableData("SAPLFSKBTABLE", "D/C", 1, "", "","Credit", False)
Call SetTableData("SAPLFSKBTABLE", "D/C", 2, "", "","Debit", False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", 1, "", "",DT_FB50L_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", 2, "", "",DT_FB50L_0100_TABLECELL_AMOUNT_IN_DOCCURR_1, False)
Call SetTableData("SAPLFSKBTABLE", "Text", 1, "", "",DT_FB50L_0100_TABLECELL_TEXT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Text", 2, "", "",DT_FB50L_0100_TABLECELL_TEXT_1, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", 2, "", "",DT_FB50L_0100_TABLECELL_PROFIT_CENTER_0, False)
Call SetTableData("SAPLFSKBTABLE", "Profit center", 1, "", "",DT_FB50L_0100_TABLECELL_PROFIT_CENTER_0, False)
Call SetTableData("SAPLFSKBTABLE", "Profit center", 2, "", "",DT_FB50L_0100_TABLECELL_PROFIT_CENTER_1, False)
Call SetTableData("SAPLFSKBTABLE", "Business area",1, "", "",DT_FB50L_0100_TABLECELL_BUSINESS_AREA_1, False)
Call TakeScreenShot
Call PressEnter()
Call PressEnter()
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetStatusBar("item1", "DT_OP_DOC_NO")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call VerifyStatusBar(DT_FB50L_1000_CHECK_TEXT_OF_COMPANY_CODE)
Call ClickButtonIfExist("Cancel   \(F12\)",False)
Wait 5

Call TakeScreenShot
Call ClickButtonIfExist("Yes",False)
'''''''--------TransactionCode-/FB03----------''''


Call SetTcode(DT_FB50L_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call SetTextbox("Document Number","RF05L-BELNR","",DT_OP_DOC_NO,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB50L_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FB50L_0100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()
Call ClickButtonIfExist("Display Document Header   \(F5\)",False)
Call TakeScreenShot
Call GetTextboxValue("BKPF-KURSF", "", "DT_OP_FB50L_1710_CHECK_TEXT_OF_EXCHANGE_RATE", True)
'Added below line to get the updated value from excel
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call VerifyTextBoxContent("Posting Date", "BKPF-BUDAT", "", ConvertDate(DT_FB50L_1710_CHECK_TEXT_OF_DOCUMENT_DATE), True)
Call VerifyTextBoxContent("Document Date", "BKPF-BLDAT", "", ConvertDate(DT_FB50L_1710_CHECK_TEXT_OF_DOCUMENT_DATE), True)
Call PressEnter()
Call VerifyGridCellContent("", 1, "Posting Key","",DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "Posting Key","",DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 1, "Account","",DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "Account","",DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 1, "Amount","",DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("", 2, "Amount","",DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("", 1, "Text","",DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
Call VerifyGridCellContent("", 2, "Text","",DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SGTXT)
Call VerifyGridCellContent("", 1, "WBS Element","",DT_FB50_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PROJK_EXT)
Call DoubleClickGuiGridCell("", 0, 1, "Account", False)
Call TakeScreenShot
Call VerifyTextBoxContent("Amt.in loc.cur.","BSEG-DMBTR","",Replace(DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMBTR,".",","), False)
Call ClickButtonIfExist("Back   \(F3\)",False)
Call DoubleClickGuiGridCell("", 0, 2, "Account", False)
Call TakeScreenShot
Call VerifyTextBoxContent("Amt.in loc.cur.","BSEG-DMBTR","",Replace(DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMBTR,".",","), False)
Call TakeScreenShot 
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
