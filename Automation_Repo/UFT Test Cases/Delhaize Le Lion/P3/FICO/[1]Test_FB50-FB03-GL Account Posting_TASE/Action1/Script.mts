'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_FB50-FB03-GL Account Posting 
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_FB50-FB03-GL Account Posting"
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

'Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''--------TransactionCode-FB50----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE) 


Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB50_1000_COMPANY_CODE,False)
Call TakeScreenShot
Call PressEnter()
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call SetTextbox("Document type","ACGL_HEAD-BLART","",DT_FB50_1010_DOCUMENT_TYPE,False)
Call SetTextbox("Currency","ACGL_HEAD-WAERS","",DT_FB50_1010_CURRENCY,False)
Call SetTextbox("Reference","ACGL_HEAD-XBLNR","",DT_FB50_1010_REFERENCE,False)
Call SetTextbox("Document Date","ACGL_HEAD-BLDAT","",ConvertDate(DT_FB50_1010_DOCUMENT_DATE),False)
Call TakeScreenShot
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", 1, "", "",DT_FB50_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "D/C", 1, "", "",DT_FB50_0100_TABLECELL_DC_0, False)
Call SetTableData("SAPLFSKBTABLE", "G/L acct", 1, "", "",DT_FB50_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", 1, "", "",DT_FB50_0100_TABLECELL_COST_CENTER_0, False)
Call PressEnter()
Call SetTableData("SAPLFSKBTABLE", "Tax code", 1, "", "",DT_FB50_0100_TABLECELL_TAX_CODE_0, False)
Call PressEnter()
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", 2, "", "",DT_FB50_0100_TABLECELL_AMOUNT_IN_DOCCURR_1, False)
Call SetTableData("SAPLFSKBTABLE", "D/C", 2, "", "",DT_FB50_0100_TABLECELL_DC_1, False)
Call SetTableData("SAPLFSKBTABLE", "G/L acct", 2, "", "",DT_FB50_0100_TABLECELL_GL_ACCT_1, False)
Call SetTableData("SAPLFSKBTABLE", "Profit center", 2, "", "",DT_FB50_0100_TABLECELL_PROFIT_CENTER_0, False)
Call SetTableData("SAPLFSKBTABLE", "Tax code", 2, "", "",DT_FB50_0100_TABLECELL_TAX_CODE_1, False)
Call TakeScreenShot
Call PressEnter()
Call ClickButtonIfExist("Simulate Document Posting   \(F9\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetStatusBar("item1", "DT_OP_EXPECTEDVALUE_OCC1")
GetRowNo = 4
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyStatusBar(DT_FB50_1000_CHECK_TEXT_OF_COMPANY_CODE)
Call ClickButtonIfExist("Cancel   \(F12\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Yes",False)

'''''''--------TransactionCode-/FB03----------''''

Call SetTcode(DT_FB50_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB50_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB50_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FB50_0100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()

CAll ClickContextButtonToolBar("&MB_VARIANT", 0)
Call SelectMenuIdToolBar("&COL0",0)
Call ClickButtonToolBar("&FIND",0)

Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB50_0841_SEARCH_TERM_OCC1,True)
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)

Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB50_0841_SEARCH_TERM,True)
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Show Selected Fields \(F7\)",True)
Call ClickButton("Transfer   \(Enter\)",True)

Call VerifyGridCellContent("", 1, "Posting Key","",DT_FB50_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "Posting Key","",DT_FB50_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 1, "Account","",DT_FB50_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "Account","",DT_FB50_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 1, "Cost Center","",DT_FB50_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOSTL)
Call VerifyGridCellContent("", 2, "Cost Center","","")
Call VerifyGridCellContent("", 1, "WBS Element","","")
Call TakeScreenShot

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
