
'''''''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'''''''.................Test Script Name :Test_09.11.01.13.01 Reverse Settlement of AuC_P1
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

gstrTestCaseName = "Test_09.11.01.13.01 Reverse Settlement of AuC_P1"
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

''''--------TransactionCode-AIST----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

Call SetTextbox("Company code","ANLA-BUKRS","",DT_AIST_0100_COMPANY_CODE,False)
Call SetTextbox("Asset","ANLA-ANLN1","",(DT_AIST_0100_ASSET),False)
Call SetTextbox("Subnumber","ANLA-ANLN2","",DT_AIST_0100_SUBNUMBER,False)
Call SetTextbox("Document Date","ANEK-BLDAT","",ConvertDate(DT_AIST_0100_DOCUMENT_DATE),False)
Call SetTextbox("Posting Date","ANEK-BUDAT","",ConvertDate(DT_AIST_0100_POSTING_DATE),False)
Call SetTextbox("Period","ANBZ-MONAT","",DT_AIST_0100_PERIOD,False)
Call SelectCheckbox("LKO74-TESTLAUF", 0, DT_AIST_0100_TEST_RUN, False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot


Call VerifyGridCellContent("Line items", 1, "WRBTR", 0, DT_AIST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WRBTR)
Call VerifyGridCellContent("Line items", 2, "WRBTR", 0, DT_AIST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_WRBTR)
Call VerifyGridCellContent("Line items", 3, "WRBTR", 0, DT_AIST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_WRBTR)
Call VerifyGridCellContent("Line items", 4, "WRBTR", 0, DT_AIST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_WRBTR)
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call SelectCheckbox("LKO74-TESTLAUF", 0, DT_AIST_0100_TEST_RUN_OCC1, False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call VerifyGridCellContent("Line items", 1, "WRBTR", 0, DT_AIST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WRBTR_OCC1)
Call VerifyGridCellContent("Line items", 2, "WRBTR", 0, DT_AIST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_WRBTR_OCC1)
Call VerifyGridCellContent("Line items", 3, "WRBTR", 0, DT_AIST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_WRBTR_OCC1)
Call VerifyGridCellContent("Line items", 4, "WRBTR", 0, DT_AIST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_WRBTR_OCC1)

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetStatusBar("item1","DT_AIST_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Asset transaction posted with document no. "&DT_AIST_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT)
Call WriteRunTimeDataToExcelGlobalSheet("DT_AIST_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_AIST_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

''''--------TransactionCode-AS03----------''''
Call SetTcode(DT_AIST_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Asset","ANLA-ANLN1","",(DT_AIST_0100_ASSET_OCC1),False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AIST_0100_SUBNUMBER_OCC1,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AIST_0100_COMPANY_CODE_OCC1,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call ClickButton("Asset values   \(Ctrl\+F1\)",False)

Call ActivateNodeGuiTree(0, "#1;#1;#1")

Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 1, "VERAENDERUNG", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 1, "JENDE", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "VERAENDERUNG", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JENDE", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE)
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR)
Call VerifyGridCellContent("Transactions", 2, "BUBTR", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR)
Call VerifyGridCellContent("Transactions", 3, "BUBTR", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BUBTR)


Call ActivateNodeGuiTree(0, "#1;#2;#1")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation", 1, "VERAENDERUNG", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG_OCC1)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation", 1, "JENDE", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE_OCC1)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation", 4, "VERAENDERUNG", 0,DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC1)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation", 4, "JENDE", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC1)
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC1)
Call VerifyGridCellContent("Transactions", 2, "BUBTR", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR_OCC1)
Call VerifyGridCellContent("Transactions", 3, "BUBTR", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BUBTR_OCC1)

Call ActivateNodeGuiTree(0, "#1;#3;#1")
Call VerifyGridCellContent("Planned values IFRS LOCL GAAP APC, depreciation", 1, "VERAENDERUNG", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG_OCC1)
Call VerifyGridCellContent("Planned values IFRS LOCL GAAP APC, depreciation", 1, "JENDE", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE_OCC1)
Call VerifyGridCellContent("Planned values IFRS LOCL GAAP APC, depreciation", 4, "VERAENDERUNG", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC1)
Call VerifyGridCellContent("Planned values IFRS LOCL GAAP APC, depreciation", 4, "JENDE", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC1)
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC1)
Call VerifyGridCellContent("Transactions", 2, "BUBTR", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR_OCC1)
Call VerifyGridCellContent("Transactions", 3, "BUBTR", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BUBTR_OCC1)


Call ActivateNodeGuiTree(0, "#1;#1;#2")
Call VerifyGridCellContent("Planned values Pre-merger STA valuation", 1, "VERAENDERUNG", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG_OCC1)
Call VerifyGridCellContent("Planned values Pre-merger STA valuation", 1, "JENDE", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE_OCC1)
Call VerifyGridCellContent("Planned values Pre-merger STA valuation", 4, "VERAENDERUNG", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC1)
Call VerifyGridCellContent("Planned values Pre-merger STA valuation", 4, "JENDE", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC1)
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC1)
Call VerifyGridCellContent("Transactions", 2, "BUBTR", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR_OCC1)
Call VerifyGridCellContent("Transactions", 3, "BUBTR", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BUBTR_OCC1)

Call DoubleClickGuiGridCell("Transactions", 0, 1,"BUBTR",False)
Call TakeScreenShot

Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot

Call SetTextbox("Asset","ANLA-ANLN1","",(DT_AIST_0100_ASSET_OCC2),False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AIST_0100_SUBNUMBER_OCC2,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AIST_0100_COMPANY_CODE_OCC2,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Asset values   \(Ctrl\+F1\)",False)
Call ActivateNodeGuiTree(0, "#1;#1;#1")

Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 1, "VERAENDERUNG", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG_OCC2)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 1, "JENDE", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE_OCC2)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "VERAENDERUNG", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC2)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JENDE", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC2)
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC2)
Call VerifyGridCellContent("Transactions", 2, "BUBTR", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR_OCC2)

Call ActivateNodeGuiTree(0, "#1;#2;#1")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation", 1, "VERAENDERUNG", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG_OCC3)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation", 1, "JENDE", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE_OCC3)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation", 4, "VERAENDERUNG", 0,DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC3)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation", 4, "JENDE", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC3)
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC3)
Call VerifyGridCellContent("Transactions", 2, "BUBTR", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR_OCC3)


Call ActivateNodeGuiTree(0, "#1;#3;#1")
Call VerifyGridCellContent("Planned values IFRS LOCL GAAP APC, depreciation", 1, "VERAENDERUNG", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG_OCC3)
Call VerifyGridCellContent("Planned values IFRS LOCL GAAP APC, depreciation", 1, "JENDE", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE_OCC3)
Call VerifyGridCellContent("Planned values IFRS LOCL GAAP APC, depreciation", 4, "VERAENDERUNG", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC3)
Call VerifyGridCellContent("Planned values IFRS LOCL GAAP APC, depreciation", 4, "JENDE", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC3)
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC3)
Call VerifyGridCellContent("Transactions", 2, "BUBTR", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR_OCC3)

Call ActivateNodeGuiTree(0, "#1;#1;#2")
Call VerifyGridCellContent("Planned values Pre-merger STA valuation", 1, "VERAENDERUNG", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG_OCC3)
Call VerifyGridCellContent("Planned values Pre-merger STA valuation", 1, "JENDE", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE_OCC3)
Call VerifyGridCellContent("Planned values Pre-merger STA valuation", 4, "VERAENDERUNG", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC3)
Call VerifyGridCellContent("Planned values Pre-merger STA valuation", 4, "JENDE", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC3)
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC3)
Call VerifyGridCellContent("Transactions", 2, "BUBTR", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR_OCC3)

Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot

Call SetTextbox("Asset","ANLA-ANLN1","",(DT_AIST_0100_ASSET_OCC3),False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AIST_0100_SUBNUMBER_OCC3,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AIST_0100_COMPANY_CODE_OCC3,False)
Call TakeScreenShot
Call PressEnter()
Call ClickButton("Asset values   \(Ctrl\+F1\)",False)
Call ActivateNodeGuiTree(0, "#1;#1;#1")

Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 1, "VERAENDERUNG", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG_OCC35)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 1, "JENDE", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE_OCC35)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "VERAENDERUNG", 0,DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC35 )
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JENDE", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC35)
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC35)
Call VerifyGridCellContent("Transactions", 2, "BUBTR", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR_OCC35)

Call ActivateNodeGuiTree(0, "#1;#2;#1")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation", 1, "VERAENDERUNG", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG_OCC36)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation", 1, "JENDE", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE_OCC36)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation", 4, "VERAENDERUNG", 0,DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC36)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation", 4, "JENDE", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC36)
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC36)
Call VerifyGridCellContent("Transactions", 2, "BUBTR", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR_OCC36)


Call ActivateNodeGuiTree(0, "#1;#3;#1")
Call VerifyGridCellContent("Planned values IFRS LOCL GAAP APC, depreciation", 1, "VERAENDERUNG", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG_OCC36)
Call VerifyGridCellContent("Planned values IFRS LOCL GAAP APC, depreciation", 1, "JENDE", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE_OCC36)
Call VerifyGridCellContent("Planned values IFRS LOCL GAAP APC, depreciation", 4, "VERAENDERUNG", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC36)
Call VerifyGridCellContent("Planned values IFRS LOCL GAAP APC, depreciation", 4, "JENDE", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC36)
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC36)
Call VerifyGridCellContent("Transactions", 2, "BUBTR", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR_OCC36)

Call ActivateNodeGuiTree(0, "#1;#1;#2")
Call VerifyGridCellContent("Planned values Pre-merger STA valuation", 1, "VERAENDERUNG", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG_OCC36)
Call VerifyGridCellContent("Planned values Pre-merger STA valuation", 1, "JENDE", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE_OCC36)
Call VerifyGridCellContent("Planned values Pre-merger STA valuation", 4, "VERAENDERUNG", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC36)
Call VerifyGridCellContent("Planned values Pre-merger STA valuation", 4, "JENDE", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC36)
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC36)
Call VerifyGridCellContent("Transactions", 2, "BUBTR", 0, DT_AIST_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR_OCC36)

Call LogOff'
Call FinalStatus()
