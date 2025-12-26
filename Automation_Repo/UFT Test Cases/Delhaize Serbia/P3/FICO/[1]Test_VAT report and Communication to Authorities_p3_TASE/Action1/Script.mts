

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_VAT report and Communication to Authorities_p3
'.................Author : TCS 
'................ Creation Date :
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

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
gstrTestCaseName = "Test_VAT report and Communication to Authorities_p3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_013_ENA_prices_are_not_higher_than_AB_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''
''--------TransactionCode-FB50----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE) 


Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_FB50_1000_COMPANY_CODE)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",true)
Call SetTextbox("Currency","ACGL_HEAD-WAERS","","RSD",False)
Call SetTextbox("Document type","ACGL_HEAD-BLART","",DT_FB50_1010_DOCUMENT_TYPE,False)
Call SetTextbox("Currency","ACGL_HEAD-WAERS","",DT_FB50_1010_CURRENCY,False)
Call SetTextbox("Reference","ACGL_HEAD-XBLNR","",DT_FB50_1010_REFERENCE,False)
Call SetTextbox("Document Date","ACGL_HEAD-BLDAT","",ConvertDate(DT_FB50_1010_DOCUMENT_DATE),False)
Call SetTextbox("Doc.Header Text","ACGL_HEAD-BKTXT","",DT_FB50_1010_DOCHEADER_TEXT,False)
Call TakeScreenShot
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", 1, "", "",DT_FB50_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "D/C", 1, "", "","Debit", False)
Call SetTableData("SAPLFSKBTABLE", "G/L acct", 1, "", "",DT_FB50_0100_TABLECELL_GL_ACCT_0, False)
'Call SetTableData("SAPLFSKBTABLE", "Tax code", 1, "", "",DT_FB50_0100_TABLECELL_TAX_CODE_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", 2, "", "",DT_FB50_0100_TABLECELL_AMOUNT_IN_DOCCURR_1, False)
Call SetTableData("SAPLFSKBTABLE", "D/C", 2, "", "","Credit", False)
Call SetTableData("SAPLFSKBTABLE", "G/L acct", 2, "", "",DT_FB50_0100_TABLECELL_GL_ACCT_1, False)
Call SetTableData("SAPLFSKBTABLE", "Tax code", 2, "", "",DT_FB50_0100_TABLECELL_TAX_CODE_1, False)
Call TakeScreenShot
Call SetTextbox("Currency","ACGL_HEAD-WAERS","","RSD",False)
Call PressEnter()
Call SelectRowGuiTableByRow("SAPLFSKBTABLE",2, False)
Call TakeScreenShot
Call ClickButtonIfExist("Tax Base for Tax Account",False)
Call TakeScreenShot
'Call ClickButtonIfExist("Continue/Confirm   \(Enter\)",True)
Call ClickButtonIfExist("Costing",True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue/Confirm   \(Enter\)",True)
Call TakeScreenShot
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
'Adapted script to handle profit center mandatory error
If  VerifyStatusBarExist(DT_PROFIT_CENTER_ERROR_MSG)  Then 
	Call ClickCellTableByRowNo("SAPLFSKBTABLE", "Status", 1, False)
	Call SetTableData("SAPLFSKBTABLE", "Profit center", 1, "", "",DT_FB50_0100_TABLECELL_PROFITCENTER_1, False)
	Call SetTableData("SAPLFSKBTABLE", "Profit center", 2, "", "",DT_FB50_0100_TABLECELL_PROFITCENTER_2, False)
	Call TakeScreenShot
	Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
End  If
Call GetStatusBar("item1", "DT_OP_DOCUMENT_CREATED")
GetRowNo = 2
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyStatusBar(DT_DOCUMENT_TO_VERIFY)
'Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB50_1000_COMPANY_CODE,True)
Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_FB50_1000_COMPANY_CODE)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",true)
'Call PressEnter()
Call SelectMenuBar("Document;Display")
Call TakeScreenShot
Call VerifyGridCellContent("",1,"Company code","", DT_FB50_0750_CHECK_TEXT_OF_COMPANY_CODE)
Call VerifyTextBoxContent("Document Number","BKPF-BELNR","",DT_FB50_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER, False)
Call VerifyGridCellContent("",1,"Account","", DT_FB50_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("",2,"Account","", DT_FB50_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("",1,"Posting Key","", DT_FB50_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",2,"Posting Key","", DT_FB50_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
