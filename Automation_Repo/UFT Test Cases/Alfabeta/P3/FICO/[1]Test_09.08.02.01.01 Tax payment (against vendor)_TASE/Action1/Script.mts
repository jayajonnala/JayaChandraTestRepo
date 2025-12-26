		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.08.02.01.01 Tax payment (against vendor)
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

gstrTestCaseName = "Test_09.08.02.01.01 Tax payment (against vendor)"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario


'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''--------------------------------F-02-----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Type","BKPF-BLART","",DT_F02_0100_TYPE,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F02_0100_DOCUMENT_DATE),False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F02_0100_COMPANY_CODE,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F02_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F02_0100_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F02_0100_DOCHEADER_TEXT,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_0100_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()

Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_0300_PSTKY,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F02_0300_TEXT,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_0300_AMOUNT,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_0300_ACCOUNT,False)
Call SetTextbox("Profit Ctrs","COBL-PRCTR","",DT_F02_0302_PROFIT_CENTER,False)
Call TakeScreenShot
Call PressEnter()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_0300_AMOUNT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F02_0302_TEXT,False)
Call TakeScreenShot
Call SelectMenuBar("Document;Simulate")
Call TakeScreenShot
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_DOC_NR_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_NR_OUTPUT&" was posted in company code GR02")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_NR_OUTPUT",DT_DOC_NR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTcode(DT_F02_0100_OKCD)     
Call PressEnter()  

'''--------------------------------fbl3n-----------------------------

Call SetTcode(DT_F02_0100_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_F02_1000_GL_ACCOUNT,False)
Call SelectRadioButton("X_AISEL","All items", False)
Call TakeScreenShot

Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call ActivateNodeGuiTree(0, "Document;Document Number")
Wait 5
Call ClickButton("%_%%DYN007_%_APP_%-VALU_PUSH",False)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_F02_1106_DOCUMENT_NUMBER,True)
Call TakeScreenShot
Call ClickButtonIfExist("Copy   \(F8\)",True)

Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_IP = year(DT_F02_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)& ConvertDoubledigit(CSTR(Month(DT_F02_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)))& ConvertDoubledigit(CSTR(Day(DT_F02_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)))
' VerifyGridCellContent(gridTitle, gridRowNumber, gridColumnName, gridIndex, expectedValue)
Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_F02_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 1, "ZUONR", 0, DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_IP)
Call VerifyGridCellContent("", 1, "BELNR", 0, DT_F02_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 1, "BLART", 0, DT_F02_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 1, "HWAER", 0, DT_F02_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER)
Call VerifyGridCellContent("", 1, "SGTXT", 0, DT_F02_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)

''''''--------TransactionCode-FBL1N ----------''''

Call SetTcode(DT_F02_0500_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call SelectRadioButton("X_AISEL","All items",False)
Call SetTextbox("Vendor account","KD_LIFNR-LOW","",DT_F02_1000_VENDOR_ACCOUNT,False)
Call TakeScreenShot

Call ClickButtonIfExist("Dynamic selections   \(Shift\+F4\)",False)
Call ClickButton("%_%%DYN012_%_APP_%-VALU_PUSH",False)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_F02_1106_DOCUMENT_NUMBER_OCC1,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_F02_0500_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",DT_F02_0500_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)
Call ClickButtonIfExist("Transfer   \(Enter\)",True)

Call VerifyGridCellContent("", 1, "BELNR", 0, DT_F02_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_OCC1)
Call VerifyGridCellContent("", 1, "BLART", 0, DT_F02_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART_OCC1)
Call VerifyGridCellContent("", 1, "BLDAT", 0, ConvertDate(DT_F02_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))
Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_F02_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB_OCC1)
Call VerifyGridCellContent("", 1, "HWAER", 0, DT_F02_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER_OCC1)
Call VerifyGridCellContent("", 1, "SGTXT", 0, DT_F02_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT_OCC1)
Call VerifyGridCellContent("", 1, "HKONT", 0, DT_F02_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)

Call LogOff'
Call FinalStatus()
