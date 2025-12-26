'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.11.01.01.01 Capitalization of AUC_AuC Assignment of Dist. P2
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

gstrTestCaseName = "Test_09.11.01.01.01 Capitalization of AUC_AuC Assignment of Dist. P2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''''--------TransactionCode-AS03----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot


Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS03_0100_COMPANY_CODE,False)
Call SetTextbox("Asset","ANLA-ANLN1","",(DT_AS03_0100_ASSET),False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AS03_0100_SUBNUMBER,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call ClickButton("Asset values   \(Ctrl\+F1\)",False)

Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JANFANG", 0, "")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "VERAENDERUNG", 0, "")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JENDE", 0, "")

Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR)
Call VerifyGridCellContent("Transactions", 2, "BUBTR", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR)

Call DoubleClickGuiGridCell("Transactions", 0, 2, "Transaction Type", False)

Call SelectMenuIdToolBar("&COL0",1)
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_AS03_0841_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",DT_AS03_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)
Call ClickButtonIfExist("Transfer   \(Enter\)",True)

Call VerifyGridCellContent("", 1, "Amount", 0, DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PSWBT)
Call VerifyGridCellContent("", 2, "Amount", 0, DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PSWBT)
Call VerifyGridCellContent("", 3, "Amount", 0, DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_PSWBT)

Call GetGridContentByTitle("", 0, "Account", 1, "DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR_OUTPUT",DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call GetGridContentByTitle("", 0, "Account", 2, "DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR_OUTPUT",DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call GetGridContentByTitle("", 0, "Account", 3, "DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR_OUTPUT",DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
'
''''--------TransactionCode-faglb03----------''''
Call SetTcode(DT_AS03_0750_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Account Number","RACCT-LOW", "0", DT_AS03_1000_ACCOUNT_NUMBER, False)
Call PressEnter()
Call TakeScreenShot

Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call TakeScreenShot
Call DoubleClickGuiGridCell("", 0, Cint(DT_AS03_0030_GRIDCELL_7_CREDIT)+1, "Credit", False)
Call TakeScreenShot

Call SelectColumnGuiGrid("", "0", "Document Number", False)
Call ClickButtonIfExist("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Document Number","%%DYN001-LOW", "0", DT_AS03_1105_DOCUMENT_NUMBER,True)
Call ClickButton("Execute   \(Enter\)",True)
'' added by ARUD from 110 to 116
Call ClickButton("Change Layout\.\.\.   \(Ctrl\+F8\)",False)
Call ClickButtonToolBar("&FIND",0)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE",0,DT_AS03_SET_SEARCH_TERM,True)
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Show Selected Fields \(F7\)",True)
Call ClickButton("Transfer   \(Enter\)",True)

''Call VerifyGridCellContent("", 1, "ZUONR", 0, DT_AS03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)
Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_AS03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)

Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call ClickButton("Choose Ledger   \(Ctrl\+F4\)",False)
Call TakeScreenShot

Call SetTextbox("Ledger","SVALD-VALUE", "0", DT_AS03_0300_LEDGER,True)

Call ClickButton("Cont\.   \(Enter\)",True)
Call TakeScreenShot

Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call GetGridContentByTitle("", 0, "Credit", Cint(DT_AS03_0030_GRIDCELL_7_CREDIT)+1, "DT_AS03_GETVALUE_CREDIT_1_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_AS03_GETVALUE_CREDIT_1_OUTPUT",DT_AS03_GETVALUE_CREDIT_1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call VerifyGridCellContent("", Cint(DT_AS03_0030_GRIDCELL_7_CREDIT)+1, "Credit", 0, DT_AS03_GETVALUE_CREDIT_1)
Call TakeScreenShot

Call DoubleClickGuiGridCell("",0, Cint(DT_AS03_0030_GRIDCELL_7_CREDIT)+1, "Credit", False)
Call TakeScreenShot

Call SelectColumnGuiGrid("", "0", "Document Number", False)
Call ClickButtonIfExist("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Document Number","%%DYN001-LOW", "0", DT_AS03_1105_DOCUMENT_NUMBER_OCC1,True)
Call ClickButton("Execute   \(Enter\)",True)

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
'''Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Asset",True)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_AS03_SET_SEARCH_TERM_OCC1,True)
Call SetComboByKey("Search Direction",0)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonIfExist("Transfer   \(Enter\)",True)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_AS03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB_OCC1)
Call VerifyGridCellContent("", 1, "Asset", 0, DT_AS03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ANLN1)

Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot

'''''--------TransactionCode-faglb03----------''''
Call SetTcode(DT_AS03_0500_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Account Number","RACCT-LOW", "0", DT_AS03_1000_ACCOUNT_NUMBER_OCC1, False)
Call PressEnter()
Call TakeScreenShot

Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

'Call GetGridContentByTitle("", 0, "Debit", Cint(DT_AS03_0030_GRIDCELL_7_DEBIT), "DT_AS03_GET_DEBIT_1_OUTPUT")
Call GetGridContentByTitle("", 0, "Debit", Cint(DT_AS03_0030_GRIDCELL_7_DEBIT)+1, "DT_AS03_GET_DEBIT_1_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_AS03_GET_DEBIT_1_OUTPUT",DT_AS03_GET_DEBIT_1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call VerifyGridCellContent("", Cint(DT_AS03_0030_GRIDCELL_7_DEBIT)+1, "Debit", 0, DT_AS03_GET_DEBIT_1)
Call TakeScreenShot

Call TakeScreenShot
'Call DoubleClickGuiGridCell("", 0, Cint(DT_AS03_0030_GRIDCELL_7_DEBIT) ,"Debit", False)
Call DoubleClickGuiGridCell("", 0, Cint(DT_AS03_0030_GRIDCELL_7_DEBIT)+1 ,"Debit", False)
Call TakeScreenShot

Call SelectColumnGuiGrid("", "0", "Document Number", False)
Call ClickButtonIfExist("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Document Number","%%DYN001-LOW", "0", DT_AS03_1105_DOCUMENT_NUMBER_OCC2,True)
Call ClickButton("Execute   \(Enter\)",True)

Call ClickButton("Change Layout\.\.\.   \(Ctrl\+F8\)",False)
Call ClickButtonToolBar("&FIND",0)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE",0,DT_AS03_SET_SEARCH_TERM,True)
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Show Selected Fields \(F7\)",True)
Call ClickButton("Transfer   \(Enter\)",True)

''Call VerifyGridCellContent("", 1, "ZUONR", 0, DT_AS03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_OCC1)
Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_AS03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB_OCC2)

Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call ClickButton("Choose Ledger   \(Ctrl\+F4\)",False)
Call TakeScreenShot

Call SetTextbox("Ledger","SVALD-VALUE", "0", DT_AS03_0300_LEDGER_OCC1,True)

Call ClickButton("Cont\.   \(Enter\)",True)
Call TakeScreenShot

Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

'Call GetGridContentByTitle("", 0, "Debit", Cint(DT_AS03_0030_GRIDCELL_7_DEBIT_OCC1), "DT_AS03_GET_DEBIT_2_OUTPUT")
Call GetGridContentByTitle("", 0, "Debit", Cint(DT_AS03_0030_GRIDCELL_7_DEBIT_OCC1)+1, "DT_AS03_GET_DEBIT_2_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_AS03_GET_DEBIT_2_OUTPUT",DT_AS03_GET_DEBIT_2)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

'Call VerifyGridCellContent("", Cint(DT_AS03_0030_GRIDCELL_7_DEBIT_OCC1), "Debit", 0, DT_AS03_GET_DEBIT_2)
Call VerifyGridCellContent("", Cint(DT_AS03_0030_GRIDCELL_7_DEBIT_OCC1)+1, "Debit", 0, DT_AS03_GET_DEBIT_2)
Call TakeScreenShot

'Call DoubleClickGuiGridCell("",0, Cint(DT_AS03_0030_GRIDCELL_7_DEBIT_OCC1), "Debit", False)
Call DoubleClickGuiGridCell("",0, Cint(DT_AS03_0030_GRIDCELL_7_DEBIT_OCC1)+1, "Debit", False)
Call TakeScreenShot

Call SelectColumnGuiGrid("", "0", "Document Number", False)
Call ClickButtonIfExist("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Document Number","%%DYN001-LOW", "0", DT_AS03_1105_DOCUMENT_NUMBER_OCC3,True)
Call ClickButton("Execute   \(Enter\)",True)

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

Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_AS03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB_OCC3)
Call VerifyGridCellContent("", 1, "Asset", 0, DT_AS03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ANLN1_OCC1)

Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot

'''''--------TransactionCode-faglb03----------''''
Call SetTcode(DT_AS03_0500_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Account Number","RACCT-LOW", "0", DT_AS03_1000_ACCOUNT_NUMBER_OCC2, False)
Call PressEnter()
Call TakeScreenShot

Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call GetGridContentByTitle("", 0, "Debit", Cint(DT_AS03_0030_GRIDCELL_7_DEBIT_OCC2)+1, "DT_AS03_GET_DEBIT_3_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_AS03_GET_DEBIT_3_OUTPUT",DT_AS03_GET_DEBIT_3)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

'Call VerifyGridCellContent("", Cint(DT_AS03_0030_GRIDCELL_7_DEBIT_OCC2), "Debit", 0, DT_AS03_GET_DEBIT_3)
Call VerifyGridCellContent("", Cint(DT_AS03_0030_GRIDCELL_7_DEBIT_OCC2)+1, "Debit", 0, DT_AS03_GET_DEBIT_3)
Call TakeScreenShot

Call TakeScreenShot
'Call DoubleClickGuiGridCell("", 0, Cint(DT_AS03_0030_GRIDCELL_7_DEBIT_OCC2), "Debit", False)
Call DoubleClickGuiGridCell("", 0, Cint(DT_AS03_0030_GRIDCELL_7_DEBIT_OCC2)+1, "Debit", False)
Call TakeScreenShot

Call SelectColumnGuiGrid("", "0", "Document Number", False)
Call ClickButtonIfExist("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Document Number","%%DYN001-LOW", "0", DT_AS03_1105_DOCUMENT_NUMBER_OCC4,True)
Call ClickButton("Execute   \(Enter\)",True)

Call ClickButton("Change Layout\.\.\.   \(Ctrl\+F8\)",False)
Call ClickButtonToolBar("&FIND",0)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE",0,DT_AS03_SET_SEARCH_TERM,True)
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Show Selected Fields \(F7\)",True)
Call ClickButton("Transfer   \(Enter\)",True)

''Call VerifyGridCellContent("", 1, "ZUONR", 0, DT_AS03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_OCC2)
'''Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_AS03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB_OCC4)
Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_AS03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB_OCC5)

Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call ClickButton("Choose Ledger   \(Ctrl\+F4\)",False)
Call TakeScreenShot

Call SetTextbox("Ledger","SVALD-VALUE", "0", DT_AS03_0300_LEDGER_OCC2,True)

Call ClickButton("Cont\.   \(Enter\)",True)
Call TakeScreenShot

Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

'Call GetGridContentByTitle("", 0, "Debit", Cint(DT_AS03_0030_GRIDCELL_7_DEBIT_OCC1), "DT_AS03_GET_DEBIT_4_OUTPUT")
Call GetGridContentByTitle("", 0, "Debit", Cint(DT_AS03_0030_GRIDCELL_7_DEBIT_OCC1)+1, "DT_AS03_GET_DEBIT_4_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_AS03_GET_DEBIT_2_OUTPUT",DT_AS03_GET_DEBIT_4)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

'Call VerifyGridCellContent("", Cint(DT_AS03_0030_GRIDCELL_7_DEBIT_OCC1), "Debit", 0, DT_AS03_GET_DEBIT_4)
Call VerifyGridCellContent("", Cint(DT_AS03_0030_GRIDCELL_7_DEBIT_OCC1)+1, "Debit", 0, DT_AS03_GET_DEBIT_4)
Call TakeScreenShot

'Call DoubleClickGuiGridCell("",0, Cint(DT_AS03_0030_GRIDCELL_7_DEBIT_OCC1), "Debit", False)
Call DoubleClickGuiGridCell("",0, Cint(DT_AS03_0030_GRIDCELL_7_DEBIT_OCC1)+1, "Debit", False)
Call TakeScreenShot

Call SelectColumnGuiGrid("", "0", "Document Number", False)
Call ClickButtonIfExist("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Document Number","%%DYN001-LOW", "0", DT_AS03_1105_DOCUMENT_NUMBER_OCC5,True)
Call ClickButton("Execute   \(Enter\)",True)

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
'''Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Asset",True)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_AS03_SET_SEARCH_TERM_OCC1,True)
Call SetComboByKey("Search Direction",0)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonIfExist("Transfer   \(Enter\)",True)
Call TakeScreenShot

''Call VerifyGridCellContent("", 1, "ANLN1", 0, DT_AS03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ANLN1_OCC2)
Call VerifyGridCellContent("", 1, "Asset", 0, DT_AS03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ANLN1_OCC2)
Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_AS03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB_OCC5)

Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot

Call LogOff'
Call FinalStatus()

