		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.06.01.04.01 Reverse AR Document
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

gstrTestCaseName = "Test_09.06.01.04.01 Reverse AR Document"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'

''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  


'''''''--------TransactionCode-FB08----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call TakeScreenShot

Call SetTextbox("Document Number","RF05A-BELNS","",DT_FB08_0105_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB08_0105_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year", "RF05A-GJAHS", "", DT_FB08_0105_FISCAL_YEAR, False)
Call SetTextbox("Reversal Reason", "UF05A-STGRD", "", ConvertDoubledigit(Cstr(DT_FB08_0105_REVERSAL_REASON)), False)
Call SetTextbox("Posting Date", "BSIS-BUDAT", "", ConvertDate(DT_FB08_0105_POSTING_DATE), False)
Call SetTextbox("Posting period","BSIS-MONAT", "",  ConvertDoubledigit(Cstr(DT_FB08_0105_POSTING_PERIOD)), False)
Call TakeScreenShot
Call TakeScreenShot
Call ClickButton("Display document before reversal   \(F5\)",False)
Call VerifyTextBoxContent("Document Date","BKPF-BLDAT", "", ConvertDate(DT_FB08_0750_CHECK_TEXT_OF_DOCUMENT_DATE), False)
Call VerifyTextBoxContent("Document Number","BKPF-BELNR", "", DT_FB08_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER, False)
Call TakeScreenShot

Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Post   \(Ctrl\+S\)",False)

Call GetStatusBar("item1","DT_DOC_NO_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_NO_OUTPUT&" was posted in company code GR02" )
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_NO_OUTPUT",DT_DOC_NO)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

''''--------TransactionCode-FB03 ----------''''

Call SetTcode(DT_FB08_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB08_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB08_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year", "RF05L-GJAHR", "", DT_FB08_0100_FISCAL_YEAR, False)
Call PressEnter()     
Call TakeScreenShot


Call VerifyTextBoxContent("Document Number", "BKPF-BELNR", "", DT_FB08_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OCC1, False)
Call VerifyTextBoxContent("Document Date", "BKPF-BLDAT", "", ConvertDate(DT_FB08_0750_CHECK_TEXT_OF_DOCUMENT_DATE_OCC1), False)

Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 3, "BSCHL", 0, DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 3, "KTONR", 0, DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
Call VerifyGridCellContent("", 1, "Amount", 0, DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PSWBT)
Call VerifyGridCellContent("", 2, "Amount", 0, DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PSWBT)
Call VerifyGridCellContent("", 3, "Amount", 0, DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_PSWBT)

''''''''''--------TransactionCode-FBL5N ----------''''

Call SetTcode(DT_FB08_0100_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_FB08_1000_CUSTOMER_ACCOUNT,False)
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_FB08_1000_COMPANY_CODE,False)
Call SelectRadioButton("X_AISEL", "All items", False)
Call TakeScreenShot

Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call ClickButton("%_%%DYN011_%_APP_%-VALU_PUSH",False)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_FB08_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_FB08_3010_TABLECELL_SINGLE_VALUE_1,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot
Call VerifyStatusBarMessageType("S")

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB08_0841_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",DT_FB08_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB08_0841_SEARCH_TERM_OCC1,True)
Call SetComboByKey("Search Direction",DT_FB08_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB08_0841_SEARCH_TERM_OCC2,True)
Call SetComboByKey("Search Direction",DT_FB08_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB75_0841_SEARCH_TERM_OCC4,True)
Call SetComboByKey("Search Direction",DT_FB08_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB75_0841_SEARCH_TERM_OCC5,True)
Call SetComboByKey("Search Direction",DT_FB08_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB75_0841_SEARCH_TERM_OCC6,True)
Call SetComboByKey("Search Direction",DT_FB08_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonIfExist("Transfer   \(Enter\)",True)

Call VerifyGridCellContent("", 1, "ZUONR", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)
Call VerifyGridCellContent("", 2, "ZUONR", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR)
Call VerifyGridCellContent("", 1, "BELNR", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 2, "BELNR", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BELNR)
Call VerifyGridCellContent("", 1, "BLART", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 2, "BLART", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BLART)
Call VerifyGridCellContent("", 1, "BLDAT", 0, ConvertDate(DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))
Call VerifyGridCellContent("", 2, "BLDAT", 0, ConvertDate(DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BLDAT))
Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 2, "DMSHB", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB)
Call VerifyGridCellContent("", 1, "AUGBL", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AUGBL)
Call VerifyGridCellContent("", 2, "AUGBL", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AUGBL)
Call VerifyGridCellContent("", 1, "SGTXT", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
Call VerifyGridCellContent("", 2, "SGTXT", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SGTXT)
Call VerifyGridCellContent("", 1, "ZTERM", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZTERM)
Call VerifyGridCellContent("", 2, "ZTERM", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZTERM)
Call VerifyGridCellContent("", 1, "PRCTR", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("", 2, "PRCTR", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)
Call VerifyGridCellContent("", 1, "HKONT", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
Call VerifyGridCellContent("", 2, "HKONT", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)
Call TakeScreenShot

Call LogOff'
Call FinalStatus()
