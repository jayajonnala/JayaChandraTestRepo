		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :F21_F32_FBL5N


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

gstrTestCaseName = "F21_F32_FBL5N"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'

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

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F21_0300_AMOUNT,False)
Call TakeScreenShot
Call ClickButton("Display Additional Data for Document Item   \(F7\)",False)
Call TakeScreenShot

Call SetTextbox("House Bank","BSEG-HBKID","",DT_F21_0330_HOUSE_BANK,False)
Call SetTextbox("/","BSEG-HKTID","",DT_F21_0330_BSEGHKTID,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F21_0330_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F21_0330_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F21_0301_AMOUNT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F21_0301_TEXT,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call ClickButton("Post   \(Ctrl\+S\)",False)

Call GetStatusBar("item1","DT_F21_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document "&DT_F21_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" was posted in company code GR02")
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet("DT_F21_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_F21_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)


''''''--------TransactionCode-F-32----------''''

Call SetTcode(DT_F21_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Company Code","BKPF-BUKRS", "", DT_F21_0131_COMPANY_CODE, False)
Call SetTextbox("Currency","BKPF-WAERS","",DT_F21_0131_CURRENCY,False)
Call SetTextbox("Account","RF05A-AGKON","",DT_F21_0131_ACCOUNT,False)
Call SelectRadioButton("RF05A-XPOS1", "Document Number", False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call SetTextbox("From","RF05A-SEL01",0,DT_F21_0731_FROM,False)
Call SetTextbox("From","RF05A-SEL01",1,DT_F21_0731_FROM_OCC1,False)
Call TakeScreenShot
Call ClickBUtton("Process Open Items   \(Shift\+F4\)",False)
Call TakeScreenShot
Call SelectTab("TS","Partial Pmt",False)
Call TakeScreenShot
Call DoubleClick
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetStatusBar("item1","DT_F21_0131_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document "&DT_F21_0131_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" was posted in company code GR02")
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet("DT_F21_0131_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_F21_0131_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)


''''''--------TransactionCode-FBL5N----------''''
'
Call SetTcode(DT_F21_0131_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_F21_1000_CUSTOMER_ACCOUNT,False)
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_F21_1000_COMPANY_CODE,False)
Call SelectRadioButton("X_AISEL", "All items", False)
Call TakeScreenShot

Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call TakeScreenShot
Call ClickButton("%_%%DYN011_%_APP_%-VALU_PUSH",False)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_F21_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_F21_3010_TABLECELL_SINGLE_VALUE_1,True)
Call SetTableData("SAPLALDBSINGLE","Single value","3","","",DT_F21_3010_TABLECELL_SINGLE_VALUE_2,True)
Call TakeScreenShot

Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)
Call TakeScreenShot
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_F21_0841_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",DT_F21_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_F21_0841_SEARCH_TERM_OCC1,True)
Call SetComboByKey("Search Direction",DT_F21_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_F21_0841_SEARCH_TERM_OCC2,True)
Call SetComboByKey("Search Direction",DT_F21_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)


Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_F21_0841_SEARCH_TERM_OCC3,True)
Call SetComboByKey("Search Direction",DT_F21_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonIfExist("Transfer   \(Enter\)",True)
Call TakeScreenShot

DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_IP = year(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)& ConvertDoubledigit(CSTR(Month(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)))&ConvertDoubledigit(CSTR(Day(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)))
Call VerifyGridCellContent("", 1, "ZUONR", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_IP)
DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR_IP = year(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR)& ConvertDoubledigit(CSTR(Month(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR)))&ConvertDoubledigit(CSTR(Day(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR)))
Call VerifyGridCellContent("", 2, "ZUONR", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR_IP)
DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_ZUONR_IP = year(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_ZUONR)& ConvertDoubledigit(CSTR(Month(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_ZUONR)))&ConvertDoubledigit(CSTR(Day(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_ZUONR)))&DT_COSTCENTER
Call VerifyGridCellContent("", 4, "ZUONR", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_ZUONR_IP)
DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_ZUONR_IP = year(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_ZUONR)& ConvertDoubledigit(CSTR(Month(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_ZUONR)))&ConvertDoubledigit(CSTR(Day(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_ZUONR)))&DT_COSTCENTER
Call VerifyGridCellContent("", 5, "ZUONR", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_ZUONR_IP)

Call VerifyGridCellContent("", 1, "BELNR", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 2, "BELNR", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BELNR)
Call VerifyGridCellContent("", 4, "BELNR", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BELNR)
Call VerifyGridCellContent("", 5, "BELNR", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BELNR)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BLART", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 2, "BLART", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BLART)
Call VerifyGridCellContent("", 4, "BLART", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BLART)
Call VerifyGridCellContent("", 5, "BLART", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BLART)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BLDAT", 0, ConvertDate(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))
Call VerifyGridCellContent("", 2, "BLDAT", 0, ConvertDate(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BLDAT))
Call VerifyGridCellContent("", 4, "BLDAT", 0, ConvertDate(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BLDAT))
Call VerifyGridCellContent("", 5, "BLDAT", 0, ConvertDate(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BLDAT))
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 2, "DMSHB", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB)
Call VerifyGridCellContent("", 3, "DMSHB", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_DMSHB)
Call VerifyGridCellContent("", 4, "DMSHB", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_DMSHB)
Call VerifyGridCellContent("", 5, "DMSHB", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_DMSHB)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "HWAER", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER)
Call VerifyGridCellContent("", 2, "HWAER", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HWAER)
Call VerifyGridCellContent("", 3, "HWAER", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_HWAER)
Call VerifyGridCellContent("", 4, "HWAER", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_HWAER)
Call VerifyGridCellContent("", 5, "HWAER", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_HWAER)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "AUGBL", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AUGBL)
Call VerifyGridCellContent("", 2, "AUGBL", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AUGBL)
Call VerifyGridCellContent("", 4, "AUGBL", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_AUGBL)
Call VerifyGridCellContent("", 5, "AUGBL", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_AUGBL)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BUDAT", 0, ConvertDate(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUDAT))
Call VerifyGridCellContent("", 2, "BUDAT", 0, ConvertDate(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUDAT))
Call VerifyGridCellContent("", 4, "BUDAT", 0, ConvertDate(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BUDAT))
Call VerifyGridCellContent("", 5, "BUDAT", 0, ConvertDate(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BUDAT))
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "HKONT", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
Call VerifyGridCellContent("", 2, "HKONT", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)
Call VerifyGridCellContent("", 4, "HKONT", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_HKONT)
Call VerifyGridCellContent("", 5, "HKONT", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_HKONT)
Call TakeScreenShot

Call LogOff'
Call FinalStatus()
