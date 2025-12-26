

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.04.01.02.02 Reverse AP Document Mass Reverse
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

gstrTestCaseName = "Test_09.04.01.02.02 Reverse AP Document Mass Reverse"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\jjonn\Desktop\TASEWork\Data\TASE_DT_09.04.01.01.01 Manage Manual Post  Direct Domestic Vendor Invoic.xls"
'strResultFolderPath = "C:\Users\jjonn\Desktop\TASEWork\Results"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'GetRowNo=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------'''''
'
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'----------------------Tcode FB65----------------------------

'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)


'Enter the details
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB65_1000_COMPANY_CODE,True) 
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot()

Call SetComboByKey("RF05A-BUSCS",DT_FB65_1100_TRANSACTN)

Call SetTextbox("Document date","INVFO-BLDAT","",ConvertDate(DT_FB65_0010_DOCUMENT_DATE),False)
Call SetTextbox("Posting Date","INVFO-BUDAT","",ConvertDate(DT_FB65_0010_POSTING_DATE),False) 
Call SetTextbox("Vendor","INVFO-ACCNT","",DT_FB65_0010_VENDOR,False) 
Call ClickButton("Back   \(F3\)",False)

Call SetTextbox("Reference","INVFO-XBLNR","",DT_FB65_0010_REFERENCE,False) 
Call PressEnter()
Call ClickBUttonifExist("No",True)
Call PressEnter()

Call SetComboByKey("RF05A-BUSCS",DT_FB65_1100_TRANSACTN_OCC1)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB65_0010_AMOUNT,False) 
Call SetComboByKey("INVFO-MWSKZ",DT_FB65_0010_TAX_AMOUNT)
Call PressEnter()

Call SelectCheckbox("INVFO-XMWST", 0, "ON", False)
Call SetTextbox("Text","INVFO-SGTXT","",DT_FB65_0010_TEXT,False)

Call SetTableData("SAPLFSKBTABLE", "G/L acct", "1", "", "", DT_FB65_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "1", "", "", DT_FB65_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "Text", "1", "", "", DT_FB65_0100_TABLECELL_TEXT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Business area", "1", "", "", DT_FB65_0100_TABLECELL_BUSINESS_AREA_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", "1", "", "", DT_FB65_0100_TABLECELL_COST_CENTER_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", "2", "", "", DT_FB65_0100_TABLECELL_COST_CENTER_1, False)
Call TakeScreenShot
Call PressEnter()

Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call VerifyTextBoxContent("Bal\.", "RF05A-AZSAL", 0, DT_FB65_1100_CHECK_TEXT_OF_BAL, False)

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call PressEnter()
Call GetStatusBar("item1","DT_DOC_1_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_1_OUTPUT&" was posted in company code GR02" )
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_1_OUTPUT",DT_DOC_1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB65_1000_COMPANY_CODE,True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)


'-----------------------------------FB65--------------------
Call SetTcode(DT_FB65_0100_OKCD) 
Call PressEnter()     ' 

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB65_1000_COMPANY_CODE,True) 
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot()
Call SelectMenuBar("Goto;Post with Reference")
Call TakeScreenShot()

Call SetTextbox("Document Number","BKPF-BELNR","",DT_FB65_0104_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB65_0104_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","BKPF-GJAHR","",Year(DT_FB65_0104_FISCAL_YEAR),False)
Call SelectCheckbox("RF05A-CPTEX", 0, "ON", False)
Call TakeScreenShot
Call PressEnter()

Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_FB65_0100_DOCUMENT_DATE),False) 
Call SetTextbox("Type","BKPF-BLART","",DT_FB65_0100_TYPE,False) 
Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_FB65_0100_POSTING_DATE),False) 
Call SetTextbox("Reference","BKPF-XBLNR","",DT_FB65_0100_REFERENCE_OCC1,False) 
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB65_0100_ACCOUNT,False) 
Call TakeScreenShot
Call PressEnter()

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call PressEnter()
Call GetTextboxValue("MESSTXT1", 0, "DT_FB65_0104_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT", True)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FB65_0104_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_FB65_0104_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call PressEnter()
Call ClickButton("Cancel   \(F12\)",True)

''''--------TransactionCode-F.80 ----------''''

Call SetTcode(DT_FB65_0104_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Company Code","BR_BUKRS-LOW","",DT_FB65_1000_COMPANY_CODE_OCC2,False)
Call SetTextbox("Fiscal Year","BR_GJAHR-LOW","",Year(DT_FB65_1000_FISCAL_YEAR),False)
Call SetTextbox("Reason for reversal","STOGRD","",DT_FB65_1000_REASON_FOR_REVERSAL,False)
Call SetTextbox("Posting Date","STODAT","",ConvertDate(DT_FB65_1000_POSTING_DATE),False)

Call ClickButtonIfExist("%_BR_BELNR_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_FB65_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_FB65_3010_TABLECELL_SINGLE_VALUE_1,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot 
Call ClickButtonIfExist("Back   \(F3\)",False)

Call SelectCheckbox("TESTLAUF", 0, "OFF", False)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot 

Call GetLabelContentByRefLabel("Document type", 224, -64,"DT_ORIGINAL_1_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_ORIGINAL_1_OUTPUT",DT_ORIGINAL_1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call GetLabelContentByRefLabel("Document type", 224, -96,"DT_ORIGINAL_2_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_ORIGINAL_2_OUTPUT",DT_ORIGINAL_2)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call  VerifyifGuiLabelExistsByRelativeid(Lcase(DT_ORIGINAL_1), "wnd\[0\]/usr/lbl\[17,8\]")
Call  VerifyifGuiLabelExistsByRelativeid(Lcase(DT_ORIGINAL_2), "wnd\[0\]/usr/lbl\[17,10\]")

'''--------TransactionCode-FBL1N ----------''''

Call SetTcode(DT_FB65_0120_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SelectRadioButton("X_AISEL","All items",False)
Call SetTextbox("Company Code","KD_BUKRS-LOW","",DT_FBL1N_1000_COMPANY_CODE,False)
Call SetTextbox("Vendor account","KD_LIFNR-LOW","",DT_FB65_1000_VENDOR_ACCOUNT,False)
Call TakeScreenShot

Call ClickButtonIfExist("Dynamic selections   \(Shift\+F4\)",False)
Call ClickButton("%_%%DYN012_%_APP_%-VALU_PUSH",False)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_FB65_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_FB65_3010_TABLECELL_SINGLE_VALUE_1,True)
Call SetTableData("SAPLALDBSINGLE","Single value","3","","",DT_CLNG_DOC1,True)
Call SetTableData("SAPLALDBSINGLE","Single value","4","","",DT_CLNG_DOC2,True)

Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot  

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB65_0841_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",DT_FB65_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonIfExist("Transfer   \(Enter\)",True)
Call TakeScreenShot

DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_IP = year(DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)& ConvertDoubledigit(CSTR(Month(DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)))& ConvertDoubledigit(CSTR(Day(DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)))

Call VerifyGridCellContent("", 1, "ICO_AUGP", 0, DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ICO_AUGP)
Call VerifyGridCellContent("", 1, "ZUONR", 0, DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_IP)
Call VerifyGridCellContent("", 1, "BELNR", 0, DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 1, "BLART", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BLART)
Call VerifyGridCellContent("", 1, "BLDAT", 0, ConvertDate(DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))
Call VerifyGridCellContent("", 1, "DMSHB", 0, CSTR(DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB))
Call VerifyGridCellContent("", 1, "HWAER", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER)
Call VerifyGridCellContent("", 1, "AUGBL", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AUGBL)
Call VerifyGridCellContent("", 1, "SGTXT", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
Call VerifyGridCellContent("", 1, "HKONT", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)

Call VerifyGridCellContent("", 2, "BELNR", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 2, "BLART", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BLART)
Call VerifyGridCellContent("", 2, "BLDAT", 0, ConvertDate(DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BLDAT))
Call VerifyGridCellContent("", 2, "DMSHB", 0, CSTR(DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB))
Call VerifyGridCellContent("", 2, "HWAER", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HWAER)
Call VerifyGridCellContent("", 2, "SGTXT", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SGTXT)
Call VerifyGridCellContent("", 2, "HKONT", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)


Call VerifyGridCellContent("", 3, "BELNR", 0, DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BELNR)
Call VerifyGridCellContent("", 3, "BLART", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BLART)
Call VerifyGridCellContent("", 3, "BLDAT", 0, ConvertDate(DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BLDAT))
Call VerifyGridCellContent("", 3, "DMSHB", 0, CSTR(DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_DMSHB))
Call VerifyGridCellContent("", 3, "HWAER", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_HWAER)
Call VerifyGridCellContent("", 3, "AUGBL", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AUGBL)
Call VerifyGridCellContent("", 3, "SGTXT", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_SGTXT)
Call VerifyGridCellContent("", 3, "HKONT", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_HKONT)

Call VerifyGridCellContent("", 4, "BELNR", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BELNR)
Call VerifyGridCellContent("", 4, "BLART", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BLART)
Call VerifyGridCellContent("", 4, "BLDAT", 0, ConvertDate(DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BLDAT))
Call VerifyGridCellContent("", 4, "DMSHB", 0, CSTR(DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_DMSHB))
Call VerifyGridCellContent("", 4, "HWAER", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_HWAER)
Call VerifyGridCellContent("", 4, "SGTXT", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_SGTXT)
Call VerifyGridCellContent("", 4, "HKONT", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_HKONT)

Call VerifyGridCellContent("", 6, "DMSHB", 0, CSTR(DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_6_DMSHB))
Call TakeScreenshot()

Call LogOff'
Call FinalStatus()

