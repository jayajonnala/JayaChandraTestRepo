
'''''''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'''''''.................Test Script Name :Test_09.11.01.13.01 Reverse Settlement of AuC_P2
'''''''.................Author : TCS 
'''''''................ Creation Date :
'''''''.................Modified By :
'''''''.................Modified Date/Details :
'''''''
'''''''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
''''''

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

gstrTestCaseName = "Test_09.11.01.13.01 Reverse Settlement of AuC_P2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'''gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


''''''''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'''''DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''''--------TransactionCode-FAGLB03----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

Call SetTextbox("Account Number","RACCT-LOW","",DT_FAGLB03_1000_ACCOUNT_NUMBER,False)
Call SetTextbox("Company Code","RBUKRS-LOW","",(DT_FAGLB03_1000_COMPANY_CODE),False)
Call SetTextbox("Fiscal Year","RYEAR","",DT_FAGLB03_1000_FISCAL_YEAR,False)

Call ClickButton("Execute   \(F8\)",False)

Call TakeScreenShot
Call DoubleClickGuiGridCell("", 0, Cint(DT_FAGLB03_0030_GRIDCELL_8_CREDIT)+1, "Credit", False)
Call TakeScreenShot

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

Call SelectColumnGuiGrid("", "0", "Asset", False)
Call ClickButtonIfExist("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Asset","%%DYN001-LOW", "0", DT_ASSET,True)
Call ClickButton("Execute   \(Enter\)",True)

Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 2, "DMSHB", 0, DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB)

Call GetGridContentByTitle("", 0, "Document Number", 1,"DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_OUTPUT",DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call GetGridContentByTitle("", 0, "Document Number", 2,"DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BELNR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BELNR_OUTPUT",DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BELNR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call ClickButton("Exit   \(Shift\+F3\)",False)
Call TakeScreenShot
Call ClickButton("Exit   \(Shift\+F3\)",False)
Call TakeScreenShot

Call SetTextbox("Account Number","RACCT-LOW","",DT_FAGLB03_1000_ACCOUNT_NUMBER_OCC1,False)
Call SetTextbox("Company Code","RBUKRS-LOW","",(DT_FAGLB03_1000_COMPANY_CODE_OCC1),False)
Call SetTextbox("Fiscal Year","RYEAR","",DT_FAGLB03_1000_FISCAL_YEAR_OCC1,False)
Call PressEnter()
Call ClickButton("Execute   \(F8\)",False)

Call TakeScreenShot
Call DoubleClickGuiGridCell("", 0, Cint(DT_FAGLB03_0030_GRIDCELL_8_DEBIT)+1, "Debit", False)
Call TakeScreenShot

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

Call SelectColumnGuiGrid("", "0", "Asset", False)
Call ClickButtonIfExist("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Asset","%%DYN001-LOW", "0", DT_ASSET_1,True)
Call ClickButton("Execute   \(Enter\)",True)

Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB_OCC1)
Call VerifyGridCellContent("", 2, "DMSHB", 0, DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB_OCC1)
Call VerifyGridCellContent("", 1, "BELNR", 0, DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_OCC1)
Call VerifyGridCellContent("", 2, "BELNR", 0, DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BELNR_OCC1)

Call ClickButton("Exit   \(Shift\+F3\)",False)
Call TakeScreenShot
Call ClickButton("Exit   \(Shift\+F3\)",False)
Call TakeScreenShot

Call SetTextbox("Account Number","RACCT-LOW","",DT_FAGLB03_1000_ACCOUNT_NUMBER_OCC2,False)
Call SetTextbox("Company Code","RBUKRS-LOW","",(DT_FAGLB03_1000_COMPANY_CODE_OCC2),False)
Call SetTextbox("Fiscal Year","RYEAR","",DT_FAGLB03_1000_FISCAL_YEAR_OCC2,False)
Call PressEnter()
Call ClickButton("Execute   \(F8\)",False)

Call TakeScreenShot
Call DoubleClickGuiGridCell("", 0, Cint(DT_FAGLB03_0030_GRIDCELL_8_DEBIT_OCC1)+1, "Debit", False)
Call TakeScreenShot

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

Call SelectColumnGuiGrid("", "0", "Asset", False)
Call ClickButtonIfExist("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Asset","%%DYN001-LOW", "0", DT_ASSET_2,True)
Call ClickButton("Execute   \(Enter\)",True)

Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB_OCC2)
Call VerifyGridCellContent("", 2, "DMSHB", 0, DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB_OCC2)
Call VerifyGridCellContent("", 1, "BELNR", 0, DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_OCC2)
Call VerifyGridCellContent("", 2, "BELNR", 0, DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BELNR_OCC2)

'''--------TransactionCode-FAGLB03----------''''
Call SetTcode(DT_FAGLB03_0500_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Account Number","RACCT-LOW","",DT_FAGLB03_1000_ACCOUNT_NUMBER_OCC3,False)
Call SetTextbox("Company Code","RBUKRS-LOW","",(DT_FAGLB03_1000_COMPANY_CODE_OCC3),False)
Call SetTextbox("Fiscal Year","RYEAR","",DT_FAGLB03_1000_FISCAL_YEAR_OCC3,False)

Call ClickButton("Choose Ledger   \(Ctrl\+F4\)",False)
Call TakeScreenShot

Call SetTextbox("Ledger","SVALD-VALUE", "0", DT_FAGLB03_0300_LEDGER,True)

Call ClickButton("Cont\.   \(Enter\)",True)
Call TakeScreenShot

Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call TakeScreenShot
Call DoubleClickGuiGridCell("", 0, Cint(DT_CELL_ROW)+1, "Credit", False)
Call TakeScreenShot

Call SelectColumnGuiGrid("", "0", "Document Number", False)
Call ClickButtonIfExist("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Document Number","%%DYN001-LOW", "0", DT_FAGLB03_1105_DOCUMENT_NUMBER,True)
Call SetTextbox("to","%%DYN001-HIGH", "0", DT_FAGLB03_1105_TO,True)
Call ClickButton("Execute   \(Enter\)",True)

Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB_OCC3)
Call VerifyGridCellContent("", 2, "DMSHB", 0, DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB_OCC3)

Call ClickButton("Exit   \(Shift\+F3\)",False)
Call TakeScreenShot
Call ClickButton("Exit   \(Shift\+F3\)",False)
Call TakeScreenShot

Call SetTextbox("Account Number","RACCT-LOW","",DT_FAGLB03_1000_ACCOUNT_NUMBER_OCC4,False)
Call PressEnter()
Call ClickButtonIfExist("Execute   \(F8\)",False)

Call TakeScreenShot
Call DoubleClickGuiGridCell("", 0, Cint(DT_FAGLB03_0030_GRIDCELL_8_DEBIT_OCC2)+1, "Debit", False)
Call TakeScreenShot

Call SelectColumnGuiGrid("", "0", "Document Number", False)
Call ClickButtonIfExist("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Document Number","%%DYN001-LOW", "0", DT_FAGLB03_1105_DOCUMENT_NUMBER_OCC1,True)
Call SetTextbox("to","%%DYN001-HIGH", "0", DT_FAGLB03_1105_TO_OCC1,True)
Call ClickButton("Execute   \(Enter\)",True)

Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB_OCC4)
Call VerifyGridCellContent("", 2, "DMSHB", 0, DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB_OCC4)

Call ClickButton("Exit   \(Shift\+F3\)",False)
Call TakeScreenShot
Call ClickButton("Exit   \(Shift\+F3\)",False)
Call TakeScreenShot

Call SetTextbox("Account Number","RACCT-LOW","",DT_FAGLB03_1000_ACCOUNT_NUMBER_OCC5,False)
Call PressEnter()
Call ClickButtonIfExist("Execute   \(F8\)",False)

Call TakeScreenShot
Call DoubleClickGuiGridCell("", 0, Cint(DT_FAGLB03_0030_GRIDCELL_8_DEBIT_OCC3)+1, "Debit", False)
Call TakeScreenShot

Call SelectColumnGuiGrid("", "0", "Document Number", False)
Call ClickButtonIfExist("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Document Number","%%DYN001-LOW", "0", DT_FAGLB03_1105_DOCUMENT_NUMBER_OCC2,True)
Call SetTextbox("to","%%DYN001-HIGH", "0", DT_FAGLB03_1105_TO_OCC2,True)
Call ClickButton("Execute   \(Enter\)",True)

Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB_OCC5)
Call VerifyGridCellContent("", 2, "DMSHB", 0, DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB_OCC5)

Call Logoff'
Call FinalStatus()
