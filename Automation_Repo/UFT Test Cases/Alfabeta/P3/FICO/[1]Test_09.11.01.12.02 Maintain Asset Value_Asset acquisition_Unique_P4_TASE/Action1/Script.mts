

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.11.01.12.02 Maintain Asset Value_Asset acquisition_Unique_P4
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

gstrTestCaseName = "Test_09.11.01.12.02 Maintain Asset Value_Asset acquisition_Unique_P4"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_02GR10_002_LocVend_w_Subrange_DSD_GR_Deliv_Note_w_Trading_Goods.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''----------------------Login----------------------------
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''''''''----------------------Tcode FAGLB03----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot    

Call SetTextbox("Account Number","RACCT-LOW","",DT_FAGLB03_1000_ACCOUNT_NUMBER,False)
Call SetTextbox("Company Code","RBUKRS-LOW","",DT_FAGLB03_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RYEAR","",DT_FAGLB03_1000_FISCAL_YEAR,False)
Call TakeScreenShot   
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot   

Call DoubleClickGuiGridCell("", "", (Cint(Month(Date)+1)), "Debit", False)

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Asset",True)
Call SetComboByKey("Search Direction",0)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonIfExist("Transfer   \(Enter\)",True)
Call TakeScreenShot

Call SelectColumnGuiGrid("", "","Asset", False)
Call TakeScreenShot
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call TakeScreenShot
Call SetTextbox("Asset","%%DYN001-LOW","",DT_ASSET_1,True)
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot

Call GetGridContentByTitle("", "", "BELNR", 1, "DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_OUTPUT",DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call VerifyGridCellContent("", 1, "DMSHB", "", DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 1, "BLDAT", "", ConvertDate(DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUDAT))
Call VerifyGridCellContent("", 1, "BUDAT", "", ConvertDate(DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))

Call ClickBUtton("Exit   \(Shift\+F3\)",False)
Call TakeScreenShot
Call ClickBUtton("Exit   \(Shift\+F3\)",False)
Call TakeScreenShot

Call ClickButton("Choose Ledger   \(Ctrl\+F4\)",False)
Call SetTextBox("Ledger","SVALD-VALUE","",DT_FAGLB03_0300_LEDGER,True)
Call TakeScreenShot
Call ClickButton("Cont\.   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call DoubleClickGuiGridCell("", "", (Cint(Month(Date))), "Debit", False)

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Asset",True)
Call SetComboByKey("Search Direction",0)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonIfExist("Transfer   \(Enter\)",True)
Call TakeScreenShot

Call SelectColumnGuiGrid("", "", "Asset", False)
Call TakeScreenShot

Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call TakeScreenShot
Call SetTextbox("Asset","%%DYN001-LOW","",DT_ASSET_1,True)
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "DMSHB", "", DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB_OCC1)
Call VerifyGridCellContent("", 1, "Asset", "", DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ANLN1)

'''''''''----------------------Tcode FAGLB03----------------------------
Call SetTcode(DT_FAGLB03_0500_OKCD)     
Call PressEnter()     
Call TakeScreenShot    

Call SetTextbox("Account Number","RACCT-LOW","",DT_FAGLB03_1000_ACCOUNT_NUMBER_OCC1,False)
Call SetTextbox("Company Code","RBUKRS-LOW","",DT_FAGLB03_1000_COMPANY_CODE_OCC1,False)
Call SetTextbox("Fiscal Year","RYEAR","",DT_FAGLB03_1000_FISCAL_YEAR_OCC1,False)
Call TakeScreenShot   
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call PressEnter() 
'Call ClickButton("Exit   \(Shift\+F3\)",False)
Call TakeScreenShot

Call ClickButton("Choose Ledger   \(Ctrl\+F4\)",False)
Call SetTextBox("Ledger","SVALD-VALUE","",DT_FAGLB03_0300_LEDGER_OCC1,True)
Call ClickButton("Cont\.   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call DoubleClickGuiGridCell("", "", (Cint(Month(Date))), "Debit", False)

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Asset",True)
Call SetComboByKey("Search Direction",0)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonIfExist("Transfer   \(Enter\)",True)
Call TakeScreenShot

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Posting Date",True)
Call SetComboByKey("Search Direction",0)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonIfExist("Transfer   \(Enter\)",True)
Call TakeScreenShot

Call SelectColumnGuiGrid("", "", "Asset", False)
Call TakeScreenShot

Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call TakeScreenShot
Call SetTextbox("Asset","%%DYN001-LOW","",DT_ASSET_2,True)
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "DMSHB", "", DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB_OCC2)
Call VerifyGridCellContent("", 1, "Asset", "", DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ANLN1_OCC1)
Call VerifyGridCellContent("", 1, "BLDAT", "", ConvertDate(DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT_OCC2))
Call VerifyGridCellContent("", 1, "BUDAT", "", ConvertDate(DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUDAT_OCC2))

Call GetGridContentByTitle("", "", "BELNR", 1, "DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_OCC1_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_OCC1_OUTPUT",DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_OCC1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)


Call ClickBUtton("Exit   \(Shift\+F3\)",False)
Call TakeScreenShot

Call ClickBUtton("Exit   \(Shift\+F3\)",False)
Call TakeScreenShot

Call ClickBUtton("Exit   \(Shift\+F3\)",False)
Call TakeScreenShot


''------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************



