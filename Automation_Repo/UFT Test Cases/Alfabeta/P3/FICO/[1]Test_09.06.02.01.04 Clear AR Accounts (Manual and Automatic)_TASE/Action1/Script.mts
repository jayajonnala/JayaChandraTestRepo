		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.06.02.01.04 Clear AR Accounts (Manual and Automatic)

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

gstrTestCaseName = "Test_09.06.02.01.04 Clear AR Accounts (Manual and Automatic)"
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

'''''''--------TransactionCode-FB30----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Posting Date","BKPF-BUDAT", "", ConvertDate(DT_F30_0122_POSTING_DATE), False)
Call SetTextbox("Document Date","BKPF-BLDAT", "", ConvertDate(DT_F30_0122_DOCUMENT_DATE), False)
Call SetTextbox("Type","BKPF-BLART", "", DT_F30_0122_TYPE, False)
Call SetTextbox("Company Code","BKPF-BUKRS","", DT_F30_0122_COMPANY_CODE, False)
Call SetTextbox("Currency/Rate","BKPF-WAERS", "", DT_F30_0122_CURRENCYRATE, False)
Call SetTextbox("Clearing text","RF05A-AUGTX", "", DT_F30_0122_CLEARING_TEXT, False)
Call SetTextbox("PstKy","RF05A-NEWBS", "", DT_F30_0122_PSTKY, False)
Call SetTextbox("Account","RF05A-NEWKO", "", DT_F30_0122_ACCOUNT, False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Text","BSEG-SGTXT", "", DT_F30_0300_TEXT, False)
Call SetTextbox("Amount","BSEG-WRBTR", "", DT_F30_0300_AMOUNT, False)
Call SetTextbox("Tax Code","BSEG-MWSKZ", "", DT_F30_0300_TAX_CODE, False)
Call TakeScreenShot
Call SetTextbox("Business Area","COBL-GSBER", "", DT_F30_1007_BUSINESS_AREA, False)
Call SetTextbox("Cost Center","COBL-KOSTL", "", DT_F30_1007_COST_CENTER, False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
Call TakeScreenShot
Call SetTextbox("Account","RF05A-AGKON", "", DT_F30_0710_ACCOUNT, False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Select All",False)
Call TakeScreenShot
Call ClickButton("Deactivate Items",False)
Call TakeScreenShot
Call ClickButton("Field content search",False)
Call TakeScreenShot
Call SelectRadioButton("RF05A-XPOS1","Document Number",True)
Call PressEnter()
Call TakeScreenShot
Call SetTextbox("From","RF05A-SEL01", "", DT_F30_0731_FROM,True)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Activate Items",False)
Call TakeScreenShot
Call ClickButton("Clearing Text\.\.\.   \(Ctrl\+F2\)",False)
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)

Call TakeScreenShot

Call GetStatusBar("item1","DT_F30_0122_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document "&DT_F30_0122_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" was posted in company code GR02")
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet("DT_F30_0122_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_F30_0122_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

''''''--------TransactionCode-FBL5N ----------''''

Call SetTcode(DT_F30_0122_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SelectRadioButton("X_CLSEL","Cleared items",False)
Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_F30_1000_CUSTOMER_ACCOUNT,False)
Call TakeScreenShot

Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call ActivateNodeGuiTree(0, "#3;#4")
Call ClickButton("%_%%DYN010_%_APP_%-VALU_PUSH",False)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_F30_1106_CLEARING_DOCUMENT,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_F30_0841_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",DT_F30_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB70_0841_SEARCH_TERM_OCC1,True)
Call SetComboByKey("Search Direction",DT_F30_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB70_0841_SEARCH_TERM_OCC2,True)
Call SetComboByKey("Search Direction",DT_F30_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB70_0841_SEARCH_TERM_OCC3,True)
Call SetComboByKey("Search Direction",DT_F30_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB70_0841_SEARCH_TERM_OCC5,True)
Call SetComboByKey("Search Direction",DT_F30_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)
Call ClickButtonIfExist("Transfer   \(Enter\)",True)

Call VerifyGridCellContent("", 1, "BELNR", 0, DT_F30_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 2, "BELNR", 0, DT_F30_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BELNR)
Call VerifyGridCellContent("", 1, "BLART", 0, DT_F30_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 2, "BLART", 0, DT_F30_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BLART)
Call VerifyGridCellContent("", 1, "BLDAT", 0, ConvertDate(DT_F30_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))
Call VerifyGridCellContent("", 2, "BLDAT", 0, ConvertDate(DT_F30_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BLDAT))
Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_F30_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 2, "DMSHB", 0, DT_F30_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB)
Call VerifyGridCellContent("", 1, "HWAER", 0, DT_F30_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER)
Call VerifyGridCellContent("", 2, "HWAER", 0, DT_F30_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HWAER)
Call VerifyGridCellContent("", 1, "AUGBL", 0, DT_F30_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AUGBL)
Call VerifyGridCellContent("", 2, "AUGBL", 0, DT_F30_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AUGBL)
Call VerifyGridCellContent("", 1, "HKONT", 0, DT_F30_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
Call VerifyGridCellContent("", 2, "HKONT", 0, DT_F30_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)
Call LogOff()
Call FinalStatus ()



