

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.11.01.09.01 Asset Acquisition with supplier_P1
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

gstrTestCaseName = "Test_09.11.01.09.01 Asset Acquisition with supplier_P1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_02GR10_002_LocVend_w_Subrange_DSD_GR_Deliv_Note_w_Trading_Goods.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'----------------------Login----------------------------
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)

Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
Call TakeScreenShot()

''''''----------------------Tcode F-90----------------------------
'
''Create Purchase Order
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Type","BKPF-BLART","",DT_F90_0100_TYPE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F90_0100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F90_0100_ACCOUNT,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F90_0100_DOCUMENT_DATE),False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_F90_0100_POSTING_DATE),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F90_0100_CURRENCYRATE,False)
Call TakeScreenShot()

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F90_0100_COMPANY_CODE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F90_0100_REFERENCE,False)
Call TakeScreenShot()
Call PressEnter() 
Call TakeScreenShot()

Call SetTextboxNoLabel("RF05A-NEWKO","",DT_F90_0302_ACCOUNT,False)
Call SetTextbox("TType","RF05A-NEWBW","",DT_F90_0302_TTYPE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F90_0302_PSTKY,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F90_0302_TEXT,False)
Call SetTextbox("Tax Amount","BSEG-WMWST","",DT_F90_0302_TAX_AMOUNT,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F90_0302_AMOUNT,False)
Call SetTextbox("Bline Date","BSEG-ZFBDT","",ConvertDate(DT_F90_0100_BLINE_DATE),False)

Call TakeScreenShot()
Call GetTextboxValue("BSEG-SGTXT",0,"DT_F90_0100_GET_TEXT_OUTPUT",False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_F90_0100_GET_TEXT_OUTPUT",DT_F90_0100_GET_TEXT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call PressEnter() 

Call SetTextbox("Quantity","BSEG-MENGE","",DT_F90_0305_QUANTITY,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F90_0305_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F90_0305_TAX_CODE,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F90_0305_TEXT,False)
Call TakeScreenShot()

Call ClickButton("All acct assignments",False)
Call SetTextbox("Reference date","COBL-BZDAT","",ConvertDate(DT_F90_0100_REFERENCE_DATE),True)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)

Call GetTextboxValue("BSEG-SGTXT",0,"DT_F90_0100_GET_TEXT_OCC1_OUTPUT",False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_F90_0100_GET_TEXT_OCC1_OUTPUT",DT_F90_0100_GET_TEXT_OCC1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call TakeScreenShot()

Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)

Call FocusTextBoxByIndex("PK  BusA Acct                               EUR   Amount        Tax amnt","RF05A-AZEI1", 0, False)
Call SendKey("{F2}")
wait (5)
Call SelectMenuBar("Document;Simulate")
Call TakeScreenShot()
Call PressEnter()

Call VerifyTextBoxContent("C","RF05A-AZSAL", 0, DT_F90_0700_CHECK_TEXT_OF_C, False)
Call VerifyTextBoxContent("PK  BusA Acct                               EUR   Amount        Tax amnt","RF05A-AZEI1", 1, DT_F90_0700_CHECK_TEXT_OF_PK, False)
Call ClickButton("Post   \(Ctrl\+S\)",False)

Call GetStatusBar("item1","DT_F90_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document "&DT_F90_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" was posted in company code GR02")
Call WriteRunTimeDataToExcelGlobalSheet("DT_F90_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_F90_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

''''''''----------------------Tcode AS03----------------------------
Call SetTcode(DT_F90_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot    

Call SetTextbox("Asset","ANLA-ANLN1","",DT_F90_0100_ASSET,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_F90_0100_SUBNUMBER,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_F90_0100_COMPANY_CODE_OCC1,False)
Call TakeScreenShot   
Call PressEnter()     
Call TakeScreenShot   

Call VerifyTextBoxContent("Quantity","ANLA-MENGE",0,DT_F90_1140_CHECK_TEXT_OF_QUANTITY,False)
Call VerifyTextBoxContent("Capitalized on","ANLA-AKTIV",0,ConvertDate(DT_F90_1142_CHECK_TEXT_OF_CAPITALIZED_ON),False)
Call VerifyTextBoxContent("First acquisition on","ANLA-ZUGDT",0,ConvertDate(DT_F90_1142_CHECK_TEXT_OF_FIRST_ACQUISITION_ON),False)

Call ClickBUtton("Asset values   \(Ctrl\+F1\)",False)

Call PressEnter()
Call SetTextbox("Fiscal year","EDIT_JAHRE","",Year(Date),False)
Call PressEnter()

Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_F90_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR)
Call VerifyGridCellContent("Transactions", 1, "BZDAT", 0, ConvertDate(DT_F90_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BZDAT_OCC1))
Call VerifyGridCellContent("Transactions", 1, "Transaction Type", 0, DT_F90_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 1, "VERAENDERUNG", 0, DT_F90_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 1, "JENDE", 0, DT_F90_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "VERAENDERUNG", 0, DT_F90_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JENDE", 0, DT_F90_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE)

Call ActivateNodeGuiTree(0, "#1;#2;#1")

Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_F90_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR)
Call VerifyGridCellContent("Transactions", 1, "BZDAT", 0, ConvertDate(DT_F90_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BZDAT_OCC1))
Call VerifyGridCellContent("Transactions", 1, "Transaction Type", 0, DT_F90_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation", 1, "VERAENDERUNG", 0, DT_F90_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation", 1, "JENDE", 0, DT_F90_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation", 4, "VERAENDERUNG", 0, DT_F90_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation", 4, "JENDE", 0, DT_F90_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE)

Call ActivateNodeGuiTree(0, "#1;#3;#1")

Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_F90_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR)
Call VerifyGridCellContent("Transactions", 1, "BZDAT", 0, ConvertDate(DT_F90_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BZDAT_OCC1))
Call VerifyGridCellContent("Transactions", 1, "Transaction Type", 0, DT_F90_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL)
Call VerifyGridCellContent("Planned values IFRS LOCL GAAP APC, depreciation", 1, "VERAENDERUNG", 0, DT_F90_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG)
Call VerifyGridCellContent("Planned values IFRS LOCL GAAP APC, depreciation", 1, "JENDE", 0, DT_F90_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE)
Call VerifyGridCellContent("Planned values IFRS LOCL GAAP APC, depreciation", 4, "VERAENDERUNG", 0, DT_F90_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG)
Call VerifyGridCellContent("Planned values IFRS LOCL GAAP APC, depreciation", 4, "JENDE", 0, DT_F90_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE)

Call ActivateNodeGuiTree(0, "#1;#1;#2")

Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_F90_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR)
Call VerifyGridCellContent("Transactions", 1, "BZDAT", 0, ConvertDate(DT_F90_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BZDAT_OCC1))
Call VerifyGridCellContent("Transactions", 1, "Transaction Type", 0, DT_F90_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL)
Call VerifyGridCellContent("Planned values Pre-merger STA valuation", 1, "VERAENDERUNG", 0, DT_F90_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG)
Call VerifyGridCellContent("Planned values Pre-merger STA valuation", 1, "JENDE", 0, DT_F90_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE)
Call VerifyGridCellContent("Planned values Pre-merger STA valuation", 4, "VERAENDERUNG", 0, DT_F90_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG)
Call VerifyGridCellContent("Planned values Pre-merger STA valuation", 4, "JENDE", 0, DT_F90_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE)

Call DoubleClickGuiGridCell("Transactions", 0, 1, "BUBTR", False)

Call VerifyTextBoxContent("Document Number","BKPF-BELNR",0,DT_F90_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER,False)



Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_F90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_F90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 3, "BSCHL", 0, DT_F90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)

Call VerifyGridCellContent("", 1, "Account", 0, DT_F90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "Account", 0, DT_F90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 3, "Account", 0, DT_F90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)

Call VerifyGridCellContent("", 1, "Tax Code", 0, DT_F90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ)
Call VerifyGridCellContent("", 2, "Tax Code", 0, DT_F90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MWSKZ)
Call VerifyGridCellContent("", 3, "Tax Code", 0, DT_F90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_MWSKZ)

Call VerifyGridCellContent("", 1, "Text", 0, DT_F90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT1)
Call VerifyGridCellContent("", 2, "Text", 0, DT_F90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SGTXT2)

Call VerifyGridCellContent("", 1, "Amount", 0, DT_F90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("", 2, "Amount", 0, DT_F90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("", 3, "Amount", 0, DT_F90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AZBET)

Call ClickButton("Exit   \(Shift\+F3\)",False)

Call ClickButton("Exit   \(Shift\+F3\)",False)

Call ClickButton("Exit   \(Shift\+F3\)",False)
'
'
''------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************




