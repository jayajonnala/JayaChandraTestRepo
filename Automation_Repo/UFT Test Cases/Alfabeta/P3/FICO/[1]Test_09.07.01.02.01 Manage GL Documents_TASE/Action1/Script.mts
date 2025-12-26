

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.07.01.02.01 Manage GL Documents
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
	GetRowNo= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
    RunTimeResultFolder= Parameter("RunTimeResultFolder")    
End If

gstrTestCaseName = "Test_09.07.01.02.01 Manage GL Documents"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'----------------------Login----------------------------
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

'''''''''--------TransactionCode-FB50----------''''
''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_FB50_0100_OKCD)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call ClickButtonIfExist("Switch Company Code   \(F7\)",False)
Call TakeScreenShot
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB50_1000_COMPANY_CODE,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SetTextbox("Document Date","ACGL_HEAD-BLDAT", "", ConvertDate(DT_FB50_1010_DOCUMENT_DATE), False)
Call PressEnter()
Call SetTextbox("Doc\.Header Text","ACGL_HEAD-BKTXT","",DT_FB50_1010_DOCHEADER_TEXT,False)
Call SetTextbox("Document type","ACGL_HEAD-BLART", "", DT_FB50_1010_DOCUMENT_TYPE, False)

Call SetTableData("SAPLFSKBTABLE", "G/L acct", "1", "", "", DT_FB50_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "G/L acct", "2", "", "", DT_FB50_0100_TABLECELL_GL_ACCT_1, False)

Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "1", "", "", DT_FB50_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "2", "", "", DT_FB50_0100_TABLECELL_AMOUNT_IN_DOCCURR_1, False)

Call SelectCellGuiTable("SAPLFSKBTABLE", "D/C", "G/L acct", DT_FB50_0100_TABLECELL_GL_ACCT_0, False)
Call SendKey("{F4}")
Call SendKey("{TAB}")

Call SelectCellGuiTable("SAPLFSKBTABLE", "D/C", "G/L acct", DT_FB50_0100_TABLECELL_GL_ACCT_1, False)
Call SendKey("{F4}")
Wait 2
Call SendKey("{DOWN}")
Call SendKey("{TAB}")

Call SetTableData("SAPLFSKBTABLE", "Profit center", "1", "", "", DT_FB50_0100_TABLECELL_PROFIT_CENTER_0, False)
Call SetTableData("SAPLFSKBTABLE", "Profit center", "2", "", "", DT_FB50_0100_TABLECELL_PROFIT_CENTER_1, False)

Call PressEnter
Call TakeScreenShot

Call ClickButtonIfExist("Simulate Document Posting   \(F9\)",False)
Call TakeScreenShot

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot

Call GetStatusBar("item1","DT_DOC_1_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_1_OUTPUT&" was posted in company code GR02")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_1_OUTPUT",DT_DOC_1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call TakeScreenShot
Call ClickButtonIfExist("Cancel   \(F12\)",True)

'''''''''--------TransactionCode-FB03----------''''
Call SetTcode(DT_FB50_0100_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB50_0100_CHECK_TEXT_OF_DOCUMENT_NUMBER,False) 

Call PressEnter()     
Call TakeScreenShot

Call SelectMenuIdToolBar("&COL0",1)
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB50_0841_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",DT_FB50_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)
Call ClickButtonIfExist("Transfer   \(Enter\)",True)

Call VerifyGridCellContent("", 1, "PRCTR", 0, DT_FB50_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_FB50_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 1, "Amount", 0, DT_FB50_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)

Call VerifyGridCellContent("", 2, "PRCTR", 0, DT_FB50_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)
Call VerifyGridCellContent("", 2, "AZBET", 0, DT_FB50_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_FB50_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)

Call Logoff'
Call FinalStatus()

