

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.04.01.01.11 Manage Manual Direct Vendor Crediting
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

gstrTestCaseName = "Test_09.04.01.01.11 Manage Manual Direct Vendor Crediting"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\jjonn\Desktop\TASEWork\Data\TASE_DT_09.04.01.01.11 Manage Manual Direct Vendor Crediting.xls"
'strResultFolderPath = "C:\Users\jjonn\Desktop\TASEWork\Results"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'GetRowNo=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''


'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''--------TransactionCode-FB60 ----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB65_1000_COMPANY_CODE,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SetComboByKey("Transactn",DT_FB65_1100_TRANSACTN)
Call TakeScreenShot

Call SetTextboxNoLabel("INVFO-ACCNT","",DT_FB65_0010_VENDOR,False)
'Call SetTextbox("Vendor","INVFO-ACCNT","",DT_FB65_0010_VENDOR,False)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB65_0010_AMOUNT,False)
Call SetTextbox("Document date", "INVFO-BLDAT", "", ConvertDate(DT_FB65_0010_DOCUMENT_DATE), False)
Call SetTextbox("Reference", "INVFO-XBLNR", "", DT_FB65_0010_REFERENCE, False)
Call TakeScreenShot

Call SelectCheckbox("INVFO-XMWST","0","ON",False)
Call TakeScreenShot

Call SetComboByKey("INVFO-MWSKZ",DT_FB65_0010_INVFOMWSKZ)
Call TakeScreenShot

Call SetTextbox("Text", "INVFO-SGTXT", "", DT_FB65_0010_TEXT, False)
Call TakeScreenShot
Call SetTableData("SAPLFSKBTABLE", "G/L acct", "1", "", "", DT_FB65_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "1", "", "", DT_FB65_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "Text", "1", "", "", DT_FB65_0100_TABLECELL_TEXT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Business area", "1", "", "", DT_FB65_0100_TABLECELL_BUSINESS_AREA_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", "1", "", "", DT_FB65_0100_TABLECELL_COST_CENTER_0, False)
Call TakeScreenShot

Call SelectTab("TS","Details",False)
Call TakeScreenShot

Call VerifyTextBoxContent("G/L", "INVFO-HKONT", "", DT_FB65_0050_CHECK_TEXT_OF_GL, False)
Call TakeScreenShot


Call SelectTab("TS","Payment",False)
Call TakeScreenShot
Call VerifyTextBoxContent("Pmnt Terms", "INVFO-ZTERM", "", DT_FB65_0020_CHECK_TEXT_OF_PAYT_TERMS, False)
Call TakeScreenShot

Call ClickButtonIfExist("Simulate Document Posting   \(F9\)",False)
Call TakeScreenShot
Call PressEnter()
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_DOC_NUMBER_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_NUMBER_OUTPUT&" was posted in company code GR02" )
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_NUMBER_OUTPUT",DT_DOC_NUMBER)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB65_1000_COMPANY_CODE,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot

'
''''--------TransactionCode-FBL1N ----------''''

Call SetTcode(DT_FB65_1100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
'
Call SelectRadioButton("X_AISEL","All items",False)
Call SetTextbox("Company Code","KD_BUKRS-LOW","",DT_FB65_1000_COMPANY_CODE_OCC2,False)
Call SetTextbox("Vendor account","KD_LIFNR-LOW","",DT_FB65_1000_VENDOR_ACCOUNT,False)
Call TakeScreenShot

Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call ClickButton("%_%%DYN012_%_APP_%-VALU_PUSH",False)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_FB65_1106_DOCUMENT_NUMBER,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot

Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB65_0841_SEARCH_TERM_OCC0,True)
Call SetComboByKey("Search Direction",DT_FB65_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot

Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB65_0841_SEARCH_TERM_OCC1,True)
Call SetComboByKey("Search Direction",DT_FB65_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonIfExist("Transfer   \(Enter\)",True)
Call TakeScreenShot


DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_IP = year(DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)& ConvertDoubledigit(CSTR(Month(DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)))& ConvertDoubledigit(CSTR(Day(DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)))

Call VerifyGridCellContent("", 1, "ICO_AUGP", 0, DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ICO_AUGP)
Call VerifyGridCellContent("", 1, "ZUONR", 0, DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_IP)
Call VerifyGridCellContent("", 1, "BELNR", 0, DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 1, "BLART", 0, DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 1, "BLDAT", 0, ConvertDate(DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))
Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 1, "HWAER", 0, DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER)
Call VerifyGridCellContent("", 1, "SGTXT", 0, DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
Call VerifyGridCellContent("", 1, "HKONT", 0, DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
Call VerifyGridCellContent("", 1, "ZTERM", 0, DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZTERM)

Call LogOff'
Call FinalStatus()
