
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AA038 Capitalization of AUC current year acquisition
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


gstrTestCaseName = "Test_AA038 Capitalization of AUC current year acquisition"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\FICO\TASE_DT_AA010- Asset Creation - create asset directly in SAP.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


''''.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath) 

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
''
''
''''--------TransactionCode-AS01----------''''
'
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Asset Class","ANLA-ANLKL","",DT_AS01_0105_ASSET_CLASS,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS01_0105_COMPANY_CODE,False)
Call TakeScreenShot
Call PressEnter() 
Call SetTextbox("Description","ANLA-TXT50","",DT_AS01_1140_DESCRIPTION,False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call TakeScreenShot
Call SetTextbox("Cost Center","ANLZ-KOSTL","",DT_AS01_1145_COST_CENTER,False)
Call PressEnter()
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Allocations", False)
Call TakeScreenShot
Call SetTextbox("Evaluation group 2","ANLA-ORD42","",DT_AS01_1160_EVALUATION_GROUP_2,False)
Call TakeScreenShot
Call PressEnter()
Call SelectTab("TABSTRIP100", "Origin", False)
Call TakeScreenShot
Call SendKey("{F4}")
Call TakeScreenShot
Call SelectTab("G_SELONETABSTRIP", "I: Vendors by Country/Company Code", True)
Call TakeScreenShot
Call SetTextbox("Name","G_SELFLD_TAB-LOW","",DT_AS01_0220_NAME,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Copy   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Create Origin long text",False)
Call TakeScreenShot
Call SendKey("{F4}")
Call TakeScreenShot


Call SetTextbox("WBS element","G_SELFLD_TAB-LOW","",DT_AS01_0220_WBS_ELEMENT,False)
''Call SetTextbox("WBS element","ANLA-POSNR","",DT_AS01_0220_WBS_ELEMENT,False)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot
Call SetFocusGuiLabel("NEW STORE-Construction","","",True)
Call TakeScreenShot
Call ClickButton("Copy   \(Enter\)",True)
Call TakeScreenShot

Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetStatusBar("item1", "DT_AS01DOCUMENT_OUTPUT")
Call VerifyStatusBar("The asset "& DT_AS01DOCUMENT_OUTPUT & " 0 is created")

'''''--------TransactionCode-F-90----------''''
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_AS01_0105_OKCD)     
Call PressEnter()     
Call CheckTCodeScreen(DT_AS01_0105_OKCD)
Call TakeScreenShot

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_AS01_0100_COMPANY_CODE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_AS01_0100_TYPE,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_AS01_0100_DOCUMENT_DATE),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_AS01_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","","Test_" & RandomNumber(10000,40000),False)
Call TakeScreenShot

Call FocusTextBox("Account","RF05A-NEWKO",False)
Call TakeScreenShot

Call SendKey("{F4}")
wait 10
Call TakeScreenShot
Call SetTextbox("Name","G_SELFLD_TAB-LOW","",DT_AS01_0220_NAME_OCC1,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Copy   \(Enter\)",True)
Call TakeScreenShot
Call PressEnter()  

Call SetTextbox("Amount","BSEG-WRBTR","",DT_AS01_0302_AMOUNT,False)
Call SetTextbox("Tax Amount","BSEG-WMWST","",DT_AS01_0302_TAX_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_AS01_0302_TAX_CODE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_AS01_0302_PSTKY,False)
''Call SetTextbox("Account","RF05A-NEWKO","",DT_AS01DOCUMENT_OUTPUT,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_AS01_0100_ASSET_OCC2,False)
Call SetTextbox("TType","RF05A-NEWBW","",DT_AS01_0302_TTYPE,False)
Call PressEnter()
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_AS01_0305_AMOUNT,False)
Call TakeScreenShot
Call ClickButtonIfExist("Display Document Overview   \(Shift\+F2\)",False)
Call TakeScreenShot

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call PressEnter()  
Call TakeScreenShot
Call GetStatusBar("item1", "DT_DOCUMENT_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOCUMENT_OUTPUT&" was posted in company code RO02")


''''''--------TransactionCode-/FB03----------''''

Call SetTcode(DT_AS01_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_AS01_0100_OKCD)

Call SetTextbox("Document Number","RF05L-BELNR","",DT_DOCUMENT_OUTPUT,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_AS01_0100_COMPANY_CODE_OCC1,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "KTONR", 0, DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 3, "KTONR", 0, DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)

''--------TransactionCode-AS01----------''''

Call SetTcode(DT_AS01_0750_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Asset Class","ANLA-ANLKL","",DT_AS01_0105_ASSET_CLASS_OCC1,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS01_0105_COMPANY_CODE_OCC1,False)
Call TakeScreenShot
Call PressEnter() 
Call SetTextbox("Description","ANLA-TXT50","",DT_AS01_1140_DESCRIPTION_OCC1,False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call TakeScreenShot
Call SetTextbox("Cost Center","ANLZ-KOSTL","",DT_AS01_1145_COST_CENTER_OCC1,False)
Call PressEnter()
Call SelectTab("TABSTRIP100", "Allocations", False)
Call TakeScreenShot
Call SetTextbox("Evaluation group 2","ANLA-ORD42","",DT_AS01_1160_EVALUATION_GROUP_2_OCC1,False)
Call TakeScreenShot
Call PressEnter()
Call SelectTab("TABSTRIP100", "Origin", False)
Call TakeScreenShot
Call SetTextbox("Vendor","ANLA-LIFNR","",DT_AS01_1181_VENDOR,False)
Call TakeScreenShot

Call ClickButton("Create Origin long text",False)
Call TakeScreenShot

Call SetTextbox("WBS element","ANLA-POSNR","",DT_AS01_1182_WBS_ELEMENT,False)

Call TakeScreenShot

Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetStatusBar("item1", "DT_AS01DOCUMENT2_OUTPUT")
Call VerifyStatusBar("The asset "& DT_AS01DOCUMENT2_OUTPUT & " 0 is created")

''''--------TransactionCode-/AIAB----------''''
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_AS01_0105_OKCD_OCC1)     
Call PressEnter()     
Call CheckTCodeScreen(DT_AS01_0105_OKCD_OCC1)
Call TakeScreenShot

Call SetTextbox("Company Code","AICOM-BUKRS","",DT_AS01_0110_COMPANY_CODE,False)
Call SetTextbox("Asset","AICOM-ANLN1","",DT_AS01_0110_ASSET,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call SelectRowGuiGridbyRowNo("",0,1,False)
Call TakeScreenShot
Call ClickButton("Enter distribution rules   \(Shift\+F6\)",False)
Call TakeScreenShot
Call SetTableData("SAPLKOBSTC_RULES","Settlement Receiver","1","","",DT_AS01_0130_TABLECELL_SETTLEMENT_RECEIVER_0,False)
Call PressEnter()     
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "AMPEL", 0, DT_AS01_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AMPEL)
Call TakeScreenShot
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot

'''--------TransactionCode-/AIBU----------''''

Call SetTcode(DT_AS01_0500_OKCD)     
Call PressEnter()     
Call CheckTCodeScreen(DT_AS01_0500_OKCD)
Call TakeScreenShot

Call SelectCheckbox("LKO74-TESTLAUF",0,DT_AS01_0100_TEST_RUN,False)
Call TakeScreenShot
Call SetTextbox("Asset Val. Date","ANEP-BZDAT","",ConvertDate(DT_AS01_0100_ASSET_VAL_DATE),False)
Call SetTextbox("Posting Date","ANEK-BUDAT","",ConvertDate(DT_AS01_0100_POSTING_DATE),False)
Call SetTextbox("Period","ANBZ-MONAT","",DT_AS01_0100_PERIOD,False)
Call SetTextbox("Document Date","ANEK-BLDAT","",ConvertDate(DT_AS01_0100_DOCUMENT_DATE_OCC1),False)
Call SetTextbox("Document type","\*KOMK1-BLART","",DT_AS01_0100_DOCUMENT_TYPE,False)
Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS01_0100_ASSET,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

''''''--------TransactionCode-/FB03----------''''

Call SetTcode(DT_AS01_0500_OKCD_OCC1)     
Call PressEnter()     
Call CheckTCodeScreen(DT_AS01_0500_OKCD_OCC1)
Call TakeScreenShot


Call SetTextbox("Document Number","RF05L-BELNR","",DT_AS01_0100_DOCUMENT_NUMBER_OCC1,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_AS01_0100_COMPANY_CODE_OCC2,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot


Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

Call VerifyGridCellContent("", 1, "KTONR", 0, DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR_OCC1)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR_OCC1)

Call VerifyGridCellContent("", 1, "AZBET", 0, DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("", 2, "AZBET", 0, DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)



'''''''--------TransactionCode-AS03----------''''
'
Call SetTcode(DT_AS01_0100_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS01_0100_ASSET_OCC2,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS01_0100_COMPANY_CODE_OCC4,False)
Call TakeScreenShot
Call ClickButton("Asset values   \(Ctrl\+F1\)",False)
Call TakeScreenShot
'
Call VerifyGridCellContent("Transactions", 1, "BZDAT", 0, ConvertDate(DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BZDAT))
Call VerifyGridCellContent("Transactions", 2, "BZDAT", 0, ConvertDate(DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BZDAT))

Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR)
Call VerifyGridCellContent("Transactions", 2, "BUBTR", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR)

Call VerifyGridCellContent("Transactions", 1, "BWASL", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL)
Call VerifyGridCellContent("Transactions", 2, "BWASL", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BWASL)

Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JANFANG", 0, FormatBlank(DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JANFANG))
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "BEZEICHNUNG", 0, FormatBlank(DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BEZEICHNUNG))
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "VERAENDERUNG", 0, FormatBlank(DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG))
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JENDE", 0,FormatBlank( DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE))

Call SelectNodeGuiTree(0,"#1;#1;#2")
Call TakeScreenShot

Call VerifyGridCellContent("Transactions", 1, "BZDAT", 0, ConvertDate(DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BZDAT_OCC1))
Call VerifyGridCellContent("Transactions", 2, "BZDAT", 0, ConvertDate(DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BZDAT_OCC1))

''Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC1)
''Call VerifyGridCellContent("Transactions", 2, "BUBTR", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR_OCC1)

Call VerifyGridCellContent("Transactions", 1, "BWASL", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL_OCC1)
Call VerifyGridCellContent("Transactions", 2, "BWASL", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BWASL_OCC1)

Call VerifyGridCellContent("Planned values Book depreciation in group currency", 4, "JANFANG", 0, FormatBlank(DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JANFANG_OCC1))
Call VerifyGridCellContent("Planned values Book depreciation in group currency", 4, "BEZEICHNUNG", 0, FormatBlank(DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BEZEICHNUNG_OCC1))
Call VerifyGridCellContent("Planned values Book depreciation in group currency", 4, "VERAENDERUNG", 0, FormatBlank(DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC1))
Call VerifyGridCellContent("Planned values Book depreciation in group currency", 4, "JENDE", 0,FormatBlank( DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC1))

Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS01_0201_ASSET,False)
Call TakeScreenShot
Call PressEnter() 
Call SelectNodeGuiTree(0,"#1;#1;#1")
Call TakeScreenShot

Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "BEZEICHNUNG", 0, FormatBlank(DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BEZEICHNUNG_OCC2))
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "VERAENDERUNG", 0, FormatBlank(DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC2))
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JENDE", 0,FormatBlank( DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC2))


Call VerifyGridCellContent("Transactions", 1, "BZDAT", 0, ConvertDate(DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BZDAT_OCC2))
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC2)
Call VerifyGridCellContent("Transactions", 1, "BWASL", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL_OCC2)

Call SelectNodeGuiTree(0,"#1;#2;#1")
Call TakeScreenShot

Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation", 4, "BEZEICHNUNG", 0, FormatBlank(DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BEZEICHNUNG_OCC3))
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation", 4, "VERAENDERUNG", 0, FormatBlank(DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC3))
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation", 4, "JENDE", 0,FormatBlank( DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC3))


Call VerifyGridCellContent("Transactions", 1, "BZDAT", 0, ConvertDate(DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BZDAT_OCC3))
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC3)
Call VerifyGridCellContent("Transactions", 1, "BWASL", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL_OCC3)

Call LogOff()

Call FinalStatus ()



