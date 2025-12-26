
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AA075 Transfer Asset Current year partial value
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


gstrTestCaseName = "Test_AA075 partial value"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\FICO\TASE_DT_AA010- Asset Creation - create asset directly in SAP.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario


''''--------------login----------------'''''

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
Call EndDateof445PeriodByDate(DT_DATE_1,"DT_ENDING_DATE_PERIOD")
Call StartDateof445PeriodByDate(DT_TODAY,"DT_STARTING_DATE_PERIOD")

If CDate(DT_ENDING_DATE_PERIOD+1)=CDate(DT_STARTING_DATE_PERIOD) Then
	Call WriteRunTimeDataToExcelGlobalSheet ("DT_REQ_DATE",DT_DATE_1)
Else
	Call WriteRunTimeDataToExcelGlobalSheet ("DT_REQ_DATE",DT_DATE_2)
End If
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'--------TransactionCode-AS01----------''''
'
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Asset Class","ANLA-ANLKL","",DT_AS01_0105_ASSET_CLASS,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS01_0105_COMPANY_CODE,False)
Call SetTextbox("Number of similar assets","RA02S-NASSETS","",DT_AS01_0105_NUMBER_OF_SIMILAR_ASSETS,False)

Call TakeScreenShot
Call PressEnter() 
Call SetTextbox("Description","ANLA-TXT50","",DT_AS01_1140_DESCRIPTION,False)
Call TakeScreenShot
Call PressEnter() 
Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call TakeScreenShot
Call SetTextbox("Cost Center","ANLZ-KOSTL","",DT_AS01_1145_COST_CENTER,False)
Call SetTextbox("Business Area","ANLZ-GSBER","",DT_AS01_1145_BUSINESS_AREA,False)
Call PressEnter()
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Allocations", False)
Call TakeScreenShot
Call SetTextbox("Evaluation group 2","ANLA-ORD42","",DT_AS01_1160_EVALUATION_GROUP_2,False)
Call SetTextbox("Evaluation group 1","ANLA-ORD41","",FormatBlank(DT_AS01_1160_EVALUATION_GROUP_1),False)

Call TakeScreenShot
Call PressEnter()
Call SelectTab("TABSTRIP100", "Origin", False)
Call TakeScreenShot
Call SetTextbox("WBS element","ANLA-POSNR","",DT_AS01_1182_WBS_ELEMENT,False)

Call SetTextbox("Vendor","ANLA-LIFNR","",DT_AS01_1181_VENDOR,False)
Call TakeScreenShot

Call SelectTab("TABSTRIP100", "General", False)
Call SelectTab("TABSTRIP100", "Origin", False)
Call SelectTab("TABSTRIP100", "Deprec. Areas", False)

Call SetTableData("SAPLAISTTC_ANLB","DKey","1","","",DT_AS01_1190_TABLECELL_DKEY_0,False)
Call PressEnter()
Call PressEnter()
Call PressEnter()

Call SelectTab("TABSTRIP100", "General", False)
Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call PressEnter()
CAll SelectCheckbox("ANLZ-XSTIL",0,DT_AS01_1145_ASSET_SHUTDOWN,False)

Call TakeScreenShot

Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call GetStatusBar("item1", "DT_AS01_0105_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("The asset "& DT_AS01_0105_CHECK_TEXT_OF_STATUSBAR_OUTPUT & " 0 is created")

'''''--------TransactionCode-F-90----------''''
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_AS01_0105_OKCD)     
Call PressEnter()     
Call CheckTCodeScreen(DT_AS01_0105_OKCD)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_AS01_0100_COMPANY_CODE,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_REQ_DATE),False)
''''Call SetTextbox("Document Date","BKPF-BLDAT","",DT_AS01_0100_DOCUMENT_DATE,False)
''Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(Date),False)
Call SetTextbox("Type","BKPF-BLART","",DT_AS01_0100_TYPE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_AS01_0100_PSTKY,False)
''Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_AS01_0100_POSTING_DATE),False)
''Call SetTextbox("Posting Date","BKPF-BUDAT","",DT_AS01_0100_POSTING_DATE,False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_REQ_DATE),False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_AS01_0100_ACCOUNT,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_AS01_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","","Test_" & RandomNumber(10000,40000),False)
Call TakeScreenShot

Call PressEnter()  
If VerifyTextBoxEnabled("PstKy","RF05A-NEWBS",0,False)= False Then
 Call PressEnter()  
	
End If
  

Call SetTextbox("TType","RF05A-NEWBW","",DT_AS01_0302_TTYPE,False) 
 Call SetTextbox("PstKy","RF05A-NEWBS","",DT_AS01_0302_PSTKY,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_AS01_0305_TEXT,False)
Call SetTextbox("Tax Amount","BSEG-WMWST","",DT_AS01_0302_TAX_AMOUNT,False) 
Call SetTextbox("Amount","BSEG-WRBTR","",DT_AS01_0302_AMOUNT,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_AS01_0302_ACCOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_AS01_0302_TAX_CODE,False)
Call SetTextbox("Bus. Area","BSEG-GSBER","",DT_AS01_0302_BUS_AREA,False)
Call TakeScreenShot
Call PressEnter()     
Call ClickButtonIfExist("Continue   \(Enter\)",True)

Call SetTextbox("Amount","BSEG-WRBTR","",DT_AS01_0305_AMOUNT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_AS01_0305_TEXT,False)
Call PressEnter()     
Call TakeScreenShot

Call ClickButtonIfExist("Display Document Overview   \(Shift\+F2\)",False)
Call VerifyTextBoxContent("C","RF05A-AZSAL",0,DT_AS01_0700_CHECK_TEXT_OF_C,False)
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
Call GetStatusBar("item1", "DT_AS01_0100_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document "&DT_AS01_0100_CHECK_TEXT_OF_STATUSBAR_OUTPUT &" was posted in company code RO02")


''''''''--------TransactionCode-AS03----------''''
''
Call SetTcode(DT_AS01_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AS01_0100_SUBNUMBER,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS01_0100_COMPANY_CODE_OCC1,False)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot

Call VerifyTextBoxContent("Capitalized on","ANLA-AKTIV","",ConvertDate(DT_REQ_DATE),False)
Call VerifyTextBoxContent("First acquisition on","ANLA-ZUGDT","",ConvertDate(DT_REQ_DATE),False)

Call ClickButton("Asset values   \(Ctrl\+F1\)",False)

'''Call VerifyGridCellContent("Transactions", 1, "BZDAT", 0, ConvertDate(DT_AS01_0100_DOCUMENT_DATE))
Call VerifyGridCellContent("Transactions", 1, "BZDAT", 0, ConvertDate(DT_REQ_DATE))
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR)
Call VerifyGridCellContent("Transactions", 1, "BWASL", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL)
Call VerifyGridCellContent("Transactions", 1, "WAERS", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WAERS)

'
Call ClickButton("Exit   \(Shift\+F3\)",False)

'--------TransactionCode-AS01----------''''

Call SetTcode(DT_AS01_0100_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_AS01_0100_OKCD_OCC1)

Call SetTextbox("Asset Class","ANLA-ANLKL","",DT_AS01_0105_ASSET_CLASS_OCC1,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS01_0105_COMPANY_CODE_OCC1,False)
Call SetTextbox("Number of similar assets","RA02S-NASSETS","",DT_AS01_0105_NUMBER_OF_SIMILAR_ASSETS_OCC1,False)

Call TakeScreenShot
Call PressEnter() 
Call SetTextbox("Description","ANLA-TXT50","",DT_AS01_1140_DESCRIPTION_OCC1,False)
Call TakeScreenShot
Call PressEnter() 
Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call TakeScreenShot
Call SetTextbox("Cost Center","ANLZ-KOSTL","",DT_AS01_1145_COST_CENTER_OCC1,False)
Call SetTextbox("Business Area","ANLZ-GSBER","",DT_AS01_1145_BUSINESS_AREA_OCC1,False)
Call PressEnter()
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Allocations", False)
Call TakeScreenShot
Call SetTextbox("Evaluation group 2","ANLA-ORD42","",DT_AS01_1160_EVALUATION_GROUP_2_OCC1,False)
Call SetTextbox("Evaluation group 1","ANLA-ORD41","",FormatBlank(DT_AS01_1160_EVALUATION_GROUP_1_OCC1),False)

Call TakeScreenShot
Call PressEnter()
Call SelectTab("TABSTRIP100", "Origin", False)
Call TakeScreenShot
Call SetTextbox("WBS element","ANLA-POSNR","",DT_AS01_1182_WBS_ELEMENT_OCC1,False)

Call SetTextbox("Vendor","ANLA-LIFNR","",DT_AS01_1181_VENDOR_OCC1,False)
Call TakeScreenShot

Call SelectTab("TABSTRIP100", "Deprec. Areas", False)

Call SetTableData("SAPLAISTTC_ANLB","DKey","1","","",DT_AS01_1190_TABLECELL_DKEY_0_OCC1,False)
Call SetTableData("SAPLAISTTC_ANLB","UseLife","1","","",DT_AS01_1190_TABLECELL_USELIFE_0,False)
Call SetTableData("SAPLAISTTC_ANLB","UseLife","2","","",DT_AS01_1190_TABLECELL_USELIFE_1,False)
Call SetTableData("SAPLAISTTC_ANLB","Prd","1","","",DT_AS01_1190_TABLECELL_PRD_0,False)
Call SetTableData("SAPLAISTTC_ANLB","Prd","2","","",DT_AS01_1190_TABLECELL_PRD_1,False)


Call PressEnter()
Call PressEnter()
Call PressEnter()
Call PressEnter()
Call PressEnter()


Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call GetStatusBar("item1", "DT_AS01_0105_CHECK_TEXT_OF_STATUSBAR_OCC2_OUTPUT")
Call VerifyStatusBar("The asset "& DT_AS01_0105_CHECK_TEXT_OF_STATUSBAR_OCC2_OUTPUT & " 0 is created")

'''--------TransactionCode-AS02----------''''

Call SetTcode(DT_AS01_1000_OKCD)     
Call PressEnter()     
Call CheckTCodeScreen(DT_AS01_1000_OKCD)

Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AS01_0100_SUBNUMBER_OCC1,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS01_0100_COMPANY_CODE_OCC2,False)
Call TakeScreenShot
Call PressEnter() 

Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call TakeScreenShot
Call SelectCheckbox("ANLZ-XSTIL",0,DT_AS01_1145_ASSET_SHUTDOWN_OCC1,False)
Call VerifyCheckBoxValue("ANLZ-XSTIL",DT_AS01_1145_CHECK_SELECTED_OF_ASSET_SHUTDOWN)

Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Yes   \(Enter\)",True)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC3)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_AS01_0100_CHECK_TEXT_OF_STATUSBAR_OCC2)
''Call VerifyStatusBar(DT_AS01_0105_CHECK_TEXT_OF_STATUSBAR_OCC3)


'''--------TransactionCode-abumn----------''''
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)


Call SetTcode(DT_AS01_0100_OKCD_OCC2)     
Call PressEnter()     
Call CheckTCodeScreen(DT_AS01_0100_OKCD_OCC2)

Call SetTextbox("Asset","RAIFP2-ANLN1","",DT_AS01_0300_ASSET,False)
Call TakeScreenShot
Call SetTextbox("Text","RAIFP2-SGTXT","",DT_AS01_0206_TEXT,False)
Call TakeScreenShot
Call SetTextbox("Existing asset","RAIFP3-ANLN1","",DT_AS01_0320_EXISTING_ASSET,False)
Call SetTextbox("Existing asset","RAIFP3-ANLN2","",DT_AS01_0320_EXISTING_ASSET_OCC1,False)
Call TakeScreenShot
''Call SetTextbox("Asset Value Date","RAIFP1-BZDAT","",ConvertDate(DT_AS01_0100_DOCUMENT_DATE),False)
''Call SetTextbox("Asset Value Date","RAIFP1-BZDAT","",DT_AS01_0100_DOCUMENT_DATE,False)
Call SetTextbox("Asset Value Date","RAIFP1-BZDAT","",ConvertDate(DT_REQ_DATE),False)
Call TakeScreenShot
''Call SetTextbox("Posting Date","RAIFP1-BUDAT","",ConvertDate(DT_AS01_0100_DOCUMENT_DATE),False)
''Call SetTextbox("Posting Date","RAIFP1-BUDAT","",DT_AS01_0100_DOCUMENT_DATE,False)
Call SetTextbox("Posting Date","RAIFP1-BUDAT","",ConvertDate(DT_REQ_DATE),False)
Call TakeScreenShot
''Call SetTextbox("Document Date","RAIFP1-BLDAT","",ConvertDate(DT_AS01_0100_DOCUMENT_DATE),False)
''Call SetTextbox("Document Date","RAIFP1-BLDAT","",DT_AS01_0100_DOCUMENT_DATE,False)
Call SetTextbox("Document Date","RAIFP1-BLDAT","",ConvertDate(DT_REQ_DATE),False)
Call TakeScreenShot

Call SelectTab("TABSTRIP100", "Additional Details", False)
Call SetTextbox("Document type","RAIFP1-BLART","",DT_AS01_0204_DOCUMENT_TYPE,False)
Call TakeScreenShot
Call SetTextbox("Posting period","RAIFP2-MONAT","",DT_AS01_0203_POSTING_PERIOD,False)
Call TakeScreenShot
Call SetTextbox("Transfer variant","RAIFP1-TRAVA","",DT_AS01_0210_TRANSFER_VARIANT,False)
Call TakeScreenShot
Call SetTextbox("Reference","RAIFP1-XBLNR","",DT_AS01_0207_REFERENCE,False)
Call TakeScreenShot


Call SelectTab("TABSTRIP100", "Partial Transfer", False)
Call SelectRadioButton("RAIFP2-XANEU","From curr\.-yr aquis\.",False)
Call SetTextbox("Amount posted","RAIFP2-ANBTR","",DT_AS01_0401_AMOUNT_POSTED,False)
Call TakeScreenShot

Call ClickButton("Save   \(Ctrl\+S\)",False)

Call GetStatusBar("item2", "DT_AS01_0100_CHECK_TEXT_OF_STATUSBAR_OCC3_OUTPUT")
Call VerifyStatusBar("Asset transaction posted with document no. "& DT_AS01_0100_COMPANY_CODE_OCC2 &" "&DT_AS01_0100_CHECK_TEXT_OF_STATUSBAR_OCC3_OUTPUT)


'''''--------TransactionCode-AS03----------''''
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_AS01_0100_OKCD_OCC3)     
Call PressEnter()     
Call CheckTCodeScreen(DT_AS01_0100_OKCD_OCC3)
Call PressEnter()     



Call ClickButton("Asset values   \(Ctrl\+F1\)",False)

Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JENDE", 0,FormatBlank( DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC1))

'''Call VerifyGridCellContent("Transactions", 2, "BZDAT", 0, ConvertDate(DT_AS01_0100_DOCUMENT_DATE))
Call VerifyGridCellContent("Transactions", 2, "BZDAT", 0, ConvertDate(DT_REQ_DATE))
Call VerifyGridCellContent("Transactions", 2, "BUBTR", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR)
Call VerifyGridCellContent("Transactions", 2, "BWASL", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BWASL)
Call VerifyGridCellContent("Transactions", 2, "WAERS", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_WAERS)

Call SetTcode(DT_AS01_0100_OKCD_OCC4)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS01_0100_ASSET,False)
Call TakeScreenShot
Call ClickButton("Asset values   \(Ctrl\+F1\)",False)


Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JENDE", 0,FormatBlank( DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC2))

''Call VerifyGridCellContent("Transactions", 1, "BZDAT", 0, ConvertDate(DT_AS01_0100_DOCUMENT_DATE))
Call VerifyGridCellContent("Transactions", 1, "BZDAT", 0, ConvertDate(DT_REQ_DATE))
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC1)
Call VerifyGridCellContent("Transactions", 1, "BWASL", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL_OCC1)
Call VerifyGridCellContent("Transactions", 1, "WAERS", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WAERS_OCC1)

Call SelectTab("IDC_TABSTRIP", "Parameters", False)
Call VerifyTextBoxContent("Useful life","AW01_DEP_PAR-NDJAR",0,DT_AS01_0304_CHECK_TEXT_OF_USEFUL_LIFE,False)
Call VerifyTextBoxNoLabelContent("AW01_DEP_PAR-NDPER",0,DT_AS01_0304_CHECK_TEXT_OF_AW01_DEP_PARNDPER,False)
Call VerifyTextBoxContent("Remaining life","AW01_DEP_PAR-RSTNDJ",0,DT_AS01_0304_CHECK_TEXT_OF_REMAINING_LIFE,False)
Call VerifyTextBoxNoLabelContent("AW01_DEP_PAR-RSTNDP",0,DT_AS01_0304_CHECK_TEXT_OF_AW01_DEP_PARRSTNDP,False)

Call LogOff()

Call FinalStatus ()



