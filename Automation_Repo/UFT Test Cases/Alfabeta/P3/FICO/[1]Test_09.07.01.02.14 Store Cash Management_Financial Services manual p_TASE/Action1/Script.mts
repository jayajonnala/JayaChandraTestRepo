		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.07.01.02.14 Store Cash Management_Financial Services manual p
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

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_09.07.01.02.14 Store Cash Management_Financial Services manual p_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''''--------TransactionCode-F-21----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Document Date","BKPF-BLDAT", "", ConvertDate(DT_F21_0100_DOCUMENT_DATE), False)
Call SetTextbox("Type","BKPF-BLART","",DT_F21_0100_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F21_0100_COMPANY_CODE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F21_0100_PSTKY,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F21_0100_CURRENCYRATE,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F21_0100_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("G/L Acc","BSEG-HKONT","",DT_F21_0301_GL_ACC,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F21_0301_AMOUNT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F21_0301_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F21_0301_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F21_0300_AMOUNT,False)
Call SetTextbox("Profit Ctrs","COBL-PRCTR","",DT_F21_0300_ProfitCTR,False)  
Call TakeScreenShot
Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
Call TakeScreenShot

Call ClickButton("Post   \(Ctrl\+S\)",False)


Call GetStatusBar("item1","DT_DOC01_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC01_OUTPUT&" was posted in company code GR02")
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet("&DT_DOC01_OUTPUT",DT_DOC01)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Document Date","BKPF-BLDAT", "", ConvertDate(DT_F21_0100_DOCUMENT_DATE_OCC1), False)
Call SetTextbox("Type","BKPF-BLART","",DT_F21_0100_TYPE_OCC1,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F21_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F21_0100_PSTKY_OCC1,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F21_0100_CURRENCYRATE_OCC1,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F21_0100_ACCOUNT_OCC1,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("G/L Acc","BSEG-HKONT","",DT_F21_0301_GL_ACC_OCC1,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F21_0301_AMOUNT_OCC1,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F21_0301_PSTKY_OCC1,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F21_0301_ACCOUNT_OCC1,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F21_0301_AMOUNT_OCC2,False)
Call TakeScreenShot
Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
Call TakeScreenShot

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)


Call GetStatusBar("item1","DT_DOC02_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC02_OUTPUT&" was posted in company code GR02")
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet("&DT_DOC02_OUTPUT",DT_DOC02)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Document Date","BKPF-BLDAT", "", ConvertDate(DT_F21_0100_DOCUMENT_DATE_OCC2), False)
Call SetTextbox("Type","BKPF-BLART","",DT_F21_0100_TYPE_OCC2,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F21_0100_COMPANY_CODE_OCC2,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F21_0100_PSTKY_OCC2,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F21_0100_CURRENCYRATE_OCC2,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F21_0100_ACCOUNT_OCC2,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F21_0300_TAX_CODE,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_F21_1007_BUSINESS_AREA,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_F21_1007_COST_CENTER,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F21_0300_AMOUNT_OCC1,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F21_0300_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F21_0300_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F21_0300_AMOUNT_OCC2,False)
Call TakeScreenShot
Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
Call TakeScreenShot

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)


Call GetStatusBar("item1","DT_DOC03_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC03_OUTPUT&" was posted in company code GR02")
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet("&DT_DOC03_OUTPUT",DT_DOC03)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

'''''''--------TransactionCode-FBL5N----------''''

Call SetTcode(DT_F21_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call SelectCheckBox("X_MERK",0,DT_F21_1000_NOTED_ITEMS,False)
Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_F21_1000_CUSTOMER_ACCOUNT,False)
Call SelectRadioButton("X_AISEL","All items",False)
Call TakeScreenShot

Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call ClickButton("%_%%DYN011_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_F21_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_F21_3010_TABLECELL_SINGLE_VALUE_1,True)
Call TakeScreenShot
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot
CAll VerifyStatusBarMessageType("S")

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB02_0841_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",DT_FB02_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonIfExist("Transfer   \(Enter\)",True)

Call VerifyGridCellContent("", 1, "BELNR", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 2, "BELNR", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BELNR)
Call VerifyGridCellContent("", 1, "BLART", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 2, "BLART", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BLART)
Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 2, "DMSHB", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB)
Call VerifyGridCellContent("", 1, "HKONT", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
Call VerifyGridCellContent("", 2, "HKONT", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)


'''''--------TransactionCode-FAGLL03----------''''
Call SetTcode(DT_F21_0500_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_F21_1000_GL_ACCOUNT,False)

Call ClickButton("Custom Selections   \(Ctrl\+F1\)",false)
Call ActivateNodeGuiTree("0","#3;#1")
Call SetSpecialTextbox("Document Number","%%DYN001-LOW","",DT_F21_0100_DOCUMENT_NUMBER,False)
Call ClickButton("Save   \(Ctrl\+S\)",False)

Call ClickButtonIfExist("Execute   \(F8\)",False)
Call VerifyStatusBarMessageType("S")
Call TakeScreenShot

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB02_0841_SEARCH_TERM_OCC1,True)
Call SetComboByKey("Search Direction",DT_FB02_0841_SEARCH_DIRCT)
Call SelectCheckbox("GS_SEARCH-EXACT_WORD", "0", "ON",True)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonIfExist("Transfer   \(Enter\)",True)

Call VerifyGridCellContent("", 1, "BELNR", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_OCC1)
Call VerifyGridCellContent("", 1, "BLART", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART_OCC1)
Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB_OCC1)
Call VerifyGridCellContent("", 1, "KONTO", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KONTO)
Call VerifyGridCellContent("", 1, "GKONT", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_GKONT)
Call LogOff'
Call FinalStatus()
