
'''''''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'''''''.................Test Script Name :Test_09.11.01.01.01 Capitalization of AUC_AuC Assignment of Dist. Rul
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

gstrTestCaseName = "Test_09.11.01.01.01 Capitalization of AUC_AuC Assignment of Dist. Rul"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'''gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


''''''''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'''''DataRowSet =2
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call StartDateof445PeriodByDate(DT_TODAY,"DT_STARTING_DATE_PERIOD")
''Call EndDateof445PeriodByDate(DT_TODAY,"DT_ENDING_DATE_PERIOD")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

'''''''--------------login----------------'''''
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

Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 1, "VERAENDERUNG", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 1, "JENDE", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "VERAENDERUNG", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JENDE", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE)
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR)
Call VerifyGridCellContent("Transactions", 1, "BZDAT", 0, ConvertDate(DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BZDAT))

Call SelectTab("IDC_TABSTRIP", "Parameters", False)
Call TakeScreenShot

Call VerifyTextBoxContent("Ord\.dep\.start date","AW01_DEP_PAR-AFABG", 0, ConvertDate(DT_AS03_0304_CHECK_TEXT_OF_ORDDEPSTART_DATE), False)

Call SelectTab("IDC_TABSTRIP", "Planned values", False)
Call TakeScreenShot

Call ActivateNodeGuiTree(0, "#1;#2;#1")

Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation", 1, "VERAENDERUNG", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG_OCC1)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation", 1, "JENDE", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE_OCC1)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation", 4, "VERAENDERUNG", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC1)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation", 4, "JENDE", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC1)
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC1)
Call VerifyGridCellContent("Transactions", 1, "BZDAT", 0, ConvertDate(DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BZDAT_OCC1))

Call SelectTab("IDC_TABSTRIP", "Parameters", False)
''Call VerifyTextBoxContent("Ord\.dep\.start date","AW01_DEP_PAR-AFABG","",DT_AS03_0304_CHECK_TEXT_OF_ORDDEPSTART_DATE,False)
Call TakeScreenShot

Call VerifyTextBoxContent("Ord\.dep\.start date","AW01_DEP_PAR-AFABG", 0, ConvertDate(DT_AS03_0304_CHECK_TEXT_OF_ORDDEPSTART_DATE_OCC1), False)

Call ActivateNodeGuiTree(0, "#1;#3;#1")

Call SelectTab("IDC_TABSTRIP", "Planned values", False)
Call TakeScreenShot

Call VerifyGridCellContent("Planned values IFRS LOCL GAAP APC, depreciation", 1, "VERAENDERUNG", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG_OCC1)
Call VerifyGridCellContent("Planned values IFRS LOCL GAAP APC, depreciation", 1, "JENDE", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE_OCC1)
Call VerifyGridCellContent("Planned values IFRS LOCL GAAP APC, depreciation", 4, "VERAENDERUNG", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC1)
Call VerifyGridCellContent("Planned values IFRS LOCL GAAP APC, depreciation", 4, "JENDE", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC1)
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC1)
Call VerifyGridCellContent("Transactions", 1, "BZDAT", 0, ConvertDate(DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BZDAT_OCC1))

Call SelectTab("IDC_TABSTRIP", "Parameters", False)
Call TakeScreenShot

Call VerifyTextBoxContent("Ord\.dep\.start date","AW01_DEP_PAR-AFABG", 0, ConvertDate(DT_AS03_0304_CHECK_TEXT_OF_ORDDEPSTART_DATE), False)

Call ActivateNodeGuiTree(0, "#1;#1;#2")

Call SelectTab("IDC_TABSTRIP", "Planned values", False)
Call TakeScreenShot

Call VerifyGridCellContent("Planned values Pre-merger STA valuation", 1, "VERAENDERUNG", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG_OCC1)
Call VerifyGridCellContent("Planned values Pre-merger STA valuation", 1, "JENDE", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE_OCC1)
Call VerifyGridCellContent("Planned values Pre-merger STA valuation", 4, "VERAENDERUNG", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC1)
Call VerifyGridCellContent("Planned values Pre-merger STA valuation", 4, "JENDE", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC1)
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC1)
Call VerifyGridCellContent("Transactions", 1, "BZDAT", 0, ConvertDate(DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BZDAT_OCC1))

Call SelectTab("IDC_TABSTRIP", "Parameters", False)
Call TakeScreenShot
 
''''--------TransactionCode-AS03----------''''

Call SetTcode(DT_AS03_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Asset","ANLA-ANLN1","",(DT_AS03_0100_ASSET_OCC1),False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AS03_0100_SUBNUMBER,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS03_0100_COMPANY_CODE,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call ClickButton("Asset values   \(Ctrl\+F1\)",False)

Call ActivateNodeGuiTree(0, "#1;#1;#1")

DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG_OCC2= Left(DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG_OCC2,1)&".000,00"
DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE_OCC2 = Left(DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE_OCC2,1)&".000,00"
DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC2 = Left(DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC2,1)&".000,00"
DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC2 = Left(DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC2,1)&".000,00"
DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC2= Left(DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC2,1)&".000,00"

Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 1, "VERAENDERUNG", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG_OCC2)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 1, "JENDE", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE_OCC2)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "VERAENDERUNG", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC2)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JENDE", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC2)
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC2)
Call VerifyGridCellContent("Transactions", 1, "BZDAT", 0, ConvertDate(DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BZDAT))

Call SelectTab("IDC_TABSTRIP", "Parameters", False)
Call TakeScreenShot

Call VerifyTextBoxContent("Ord\.dep\.start date","AW01_DEP_PAR-AFABG", 0, ConvertDate(DT_AS03_0304_CHECK_TEXT_OF_ORDDEPSTART_DATE_OCC2), False)

Call SelectTab("IDC_TABSTRIP", "Planned values", False)
Call TakeScreenShot

Call ActivateNodeGuiTree(0, "#1;#2;#1")

Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation", 1, "VERAENDERUNG", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG_OCC2)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation", 1, "JENDE", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE_OCC2)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation", 4, "VERAENDERUNG", 0,DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC2)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation", 4, "JENDE", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC2)
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC2)
Call VerifyGridCellContent("Transactions", 1, "BZDAT", 0, ConvertDate(DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BZDAT_OCC1))

Call SelectTab("IDC_TABSTRIP", "Parameters", False)
Call TakeScreenShot

Call VerifyTextBoxContent("Ord\.dep\.start date","AW01_DEP_PAR-AFABG", 0, ConvertDate(DT_AS03_0304_CHECK_TEXT_OF_ORDDEPSTART_DATE_OCC3), False)
'
Call ActivateNodeGuiTree(0, "#1;#3;#1")

Call SelectTab("IDC_TABSTRIP", "Planned values", False)
Call TakeScreenShot

Call VerifyGridCellContent("Planned values IFRS LOCL GAAP APC, depreciation", 1, "VERAENDERUNG", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG_OCC2)
Call VerifyGridCellContent("Planned values IFRS LOCL GAAP APC, depreciation", 1, "JENDE", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE_OCC2)
Call VerifyGridCellContent("Planned values IFRS LOCL GAAP APC, depreciation", 4, "VERAENDERUNG", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC2)
Call VerifyGridCellContent("Planned values IFRS LOCL GAAP APC, depreciation", 4, "JENDE", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC2)
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC2)
Call VerifyGridCellContent("Transactions", 1, "BZDAT", 0, ConvertDate(DT_CHECK_22_16))

Call SelectTab("IDC_TABSTRIP", "Parameters", False)
Call VerifyTextBoxContent("Ord\.dep\.start date","AW01_DEP_PAR-AFABG","",ConvertDate(DT_AS03_0304_CHECK_TEXT_OF_ORDDEPSTART_DATE_OCC2),False)
Call TakeScreenShot

Call VerifyTextBoxContent("Ord\.dep\.start date","AW01_DEP_PAR-AFABG", 0, ConvertDate(DT_CHECK_22_17), False)


Call ActivateNodeGuiTree(0, "#1;#1;#2")

Call SelectTab("IDC_TABSTRIP", "Planned values", False)
Call TakeScreenShot

Call VerifyGridCellContent("Planned values Pre-merger STA valuation", 1, "VERAENDERUNG", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG_OCC2)
Call VerifyGridCellContent("Planned values Pre-merger STA valuation", 1, "JENDE", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE_OCC2)
Call VerifyGridCellContent("Planned values Pre-merger STA valuation", 4, "VERAENDERUNG", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC2)
Call VerifyGridCellContent("Planned values Pre-merger STA valuation", 4, "JENDE", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC2)
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC2)
Call VerifyGridCellContent("Transactions", 1, "BZDAT", 0, ConvertDate(DT_CHECK_20_16))

Call SelectTab("IDC_TABSTRIP", "Parameters", False)
Call TakeScreenShot

Call VerifyTextBoxContent("Ord\.dep\.start date","AW01_DEP_PAR-AFABG", 0, ConvertDate(DT_CHECK_20_17), False)
Call logoff'
Call FinalStatus()
