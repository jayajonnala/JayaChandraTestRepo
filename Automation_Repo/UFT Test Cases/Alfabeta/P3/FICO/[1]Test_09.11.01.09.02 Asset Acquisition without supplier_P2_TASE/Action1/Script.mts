		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.11.01.09.02 Asset Acquisition without supplier_P2
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

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_09.11.01.09.02 Asset Acquisition without supplier_P2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario


'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''--------TransactionCode-AS01----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

Call SetTextbox("Account Number","RACCT-LOW","",DT_FAGLB03_1000_ACCOUNT_NUMBER,False)
Call SetTextbox("Company Code","RBUKRS-LOW","",DT_FAGLB03_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RYEAR","",DT_FAGLB03_1000_FISCAL_YEAR,False)
Call TakeScreenShot
Call ClickButton("Choose Ledger   \(Ctrl\+F4\)",False)
Call TakeScreenShot
Call SetTextbox("Ledger","SVALD-VALUE","",DT_FAGLB03_0300_LEDGER,True)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call DoubleClickGuiGridCell("",0,(Cint(DT_FAGLB03_0030_GRIDCELL_6_CREDIT)+1),"Credit",False)
Call TakeScreenShot

Call ClickLabel("DocumentNo", "1", False)
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FAGLB03_1105_DOCUMENT_NUMBER,True)
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenshot

Call CLickButton("Display Document   \(Ctrl\+Shift\+F7\)",False)
Call TakeScreenshot
Call ClickButton("Call Up Document Overview   \(F9\)",False)

Call SelectMenuIdToolBar("&COL0",1)
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FAGLB03_0841_SEARCH_TERM_OCC2,True)
Call SetComboByKey("Search Direction",DT_FAGLB03_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FAGLB03_0841_SEARCH_TERM_OCC3,True)
Call SetComboByKey("Search Direction",DT_FAGLB03_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonIfExist("Transfer   \(Enter\)",True)

DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_IP = year(DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)& ConvertDoubledigit(CSTR(Month(DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)))&ConvertDoubledigit(CSTR(Day(DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)))
Call VerifyGridCellContent("", 1, "Assignment", 0, DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_IP)
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 1, "PRCTR", 0, DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("", 1, "Account Type", 0, DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_GKART)

Call ClickButton("Back   \(F3\)",False)
Wait 2
Call ClickButton("Back   \(F3\)",False)
Wait 2
Call ClickButton("Back   \(F3\)",False)
Wait 2
Call ClickButton("Choose Ledger   \(Ctrl\+F4\)",False)
Wait 2
Call SetTextbox("Ledger","SVALD-VALUE","",DT_FAGLB03_0300_LEDGER_OCC1,True)
Call PressEnter()     
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call DoubleClickGuiGridCell("",0,(Cint(DT_FAGLB03_0030_GRIDCELL_6_CREDIT_OCC1)+1),"Credit",False)
Call TakeScreenShot

Call ClickLabel("DocumentNo", "1", False)
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FAGLB03_1105_DOCUMENT_NUMBER_OCC1,True)
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenshot

Call CLickButton("Display Document   \(Ctrl\+Shift\+F7\)",False)
Call TakeScreenshot
Call ClickButton("Call Up Document Overview   \(F9\)",False)
Call TakeScreenshot

Call SelectMenuIdToolBar("&COL0",1)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FAGLB03_0841_SEARCH_TERM_OCC2,True)
Call SetComboByKey("Search Direction",DT_FAGLB03_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FAGLB03_0841_SEARCH_TERM_OCC3,True)
Call SetComboByKey("Search Direction",DT_FAGLB03_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonIfExist("Transfer   \(Enter\)",True)

DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_OCC1_IP = year(DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT_OCC1)& ConvertDoubledigit(CSTR(Month(DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT_OCC1)))&ConvertDoubledigit(CSTR(Day(DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_OCC1)))
Call VerifyGridCellContent("", 1, "Assignment", 0, DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_OCC1_IP)
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OCC1)
Call VerifyGridCellContent("", 1, "PRCTR", 0, DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR_OCC1)
Call VerifyGridCellContent("", 1, "Account Type", 0, DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_GKART_OCC1)

Call LogOff
Call FInalStatus()''


