'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Manual Posting of GL Document both ledgers_p7
'.................Author : TCS        
'................ Creation Date    : 
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


gstrTestCaseName = "Test_Manual Posting of GL Document both ledgers_p7_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Clear AP Accounts - Manual_p4_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''----------------------Login----------------------------
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''''--------TransactionCode-FB50 ----------''''
'
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INC_VAL",Cint(DT_INC_VAL)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB50_1000_COMPANY_CODE,True)
Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_FB50_1000_COMPANY_CODE)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SetTextbox("Document date", "ACGL_HEAD-BLDAT", "", ConvertDate(DT_FB50_1010_DOCUMENT_DATE), False)
Call SetTextbox("Posting Date", "ACGL_HEAD-BUDAT", "", ConvertDate(DT_FB50_1010_POSTING_DATE), False)
Call SetTextbox("Document type", "ACGL_HEAD-BLART", "", DT_FB50_1010_DOCUMENT_TYPE, False)
Call SetTextbox("Doc\.Header Text","ACGL_HEAD-BKTXT","",DT_FB50_1010_DOCHEADER_TEXT,False)
Call SetTextbox("Reference", "ACGL_HEAD-XBLNR", "", DT_FB50_1010_REFERENCE, False)

Call SetTableData("SAPLFSKBTABLE", "G/L acct", "1", "", "", DT_FB50_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "G/L acct", "2", "", "", DT_FB50_0100_TABLECELL_GL_ACCT_1, False)

Call SelectCellGuiTable("SAPLFSKBTABLE", "D/C", "G/L acct", DT_FB50_0100_TABLECELL_GL_ACCT_0, False)
Call SendKey("{F4}")
Wait 2
Call SendKey("{DOWN}")
Call SendKey("{TAB}")

Call SelectCellGuiTable("SAPLFSKBTABLE", "D/C", "G/L acct", DT_FB50_0100_TABLECELL_GL_ACCT_1, False)
Call SendKey("{F4}")
Call SendKey("{TAB}")

Call SetTableData("SAPLFSKBTABLE", "D/C", "1", "", "", DT_FB50_0100_TABLECELL_DC_0, False)
Call SetTableData("SAPLFSKBTABLE", "D/C", "2", "", "", DT_FB50_0100_TABLECELL_DC_1, False)

Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "1", "", "", DT_FB50_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "2", "", "", DT_FB50_0100_TABLECELL_AMOUNT_IN_DOCCURR_1, False)

Call SetTableData("SAPLFSKBTABLE", "Company code", "1", "", "", DT_FB50_0100_TABLECELL_COMPANY_CODE_0, False)
Call SetTableData("SAPLFSKBTABLE", "Company code", "2", "", "", DT_FB50_0100_TABLECELL_COMPANY_CODE_1, False)

Call SetTableData("SAPLFSKBTABLE", "Business area", "1", "", "", DT_FB50_0100_TABLECELL_BUSINESS_AREA_0, False)
Call SetTableData("SAPLFSKBTABLE", "Business area", "2", "", "", DT_FB50_0100_TABLECELL_BUSINESS_AREA_1, False)

Call SetTableData("SAPLFSKBTABLE", "Trans.type", "1", "", "", DT_FB50_0100_TABLECELL_TransactionType_0, False)

Call SetTableData("SAPLFSKBTABLE", "Cost center", "1", "", "", DT_FB50_0100_TABLECELL_COST_CENTER_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", "2", "", "", DT_FB50_0100_TABLECELL_COST_CENTER_1, False)

Call ClickButton("Simulate Document Posting   \(F9\)",False)
Call TakeScreenShot

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot

Call GetStatusBar("item1","DT_GETDOCUMENT_OUTPUT")
Call VerifyStatusBar("Document "&DT_GETDOCUMENT_OUTPUT&" was posted in company code RS01" )
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet ("DT_GETDOCUMENT_OUTPUT",DT_GETDOCUMENT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call LogOff()
