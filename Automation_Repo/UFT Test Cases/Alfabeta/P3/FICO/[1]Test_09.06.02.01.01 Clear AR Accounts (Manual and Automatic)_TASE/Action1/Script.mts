		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.06.02.01.01 Clear AR Accounts (Manual and Automatic)

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


gstrTestCaseName = "Test_09.06.02.01.01 Clear AR Accounts (Manual and Automatic)"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'GetRowNo =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'

'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''''''--------TransactionCode-F-32----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Company Code","BKPF-BUKRS", "", DT_F32_0131_COMPANY_CODE, False)
Call SetTextbox("Period","BKPF-MONAT","",ConvertDoubledigit(Cstr(Month(DT_F32_0131_PERIOD))),False)
Call SetTextbox("Currency","BKPF-WAERS","",DT_F32_0131_CURRENCY,False)
Call SetTextbox("Account","RF05A-AGKON","",DT_F32_0131_ACCOUNT,False)
Call SelectRadioButton("RF05A-XPOS1", "Document Number", False)
Call TakeScreenShot
Call PressEnter()  
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("From","RF05A-SEL01",0,DT_F32_0731_FROM,False)
Call SetTextbox("From","RF05A-SEL01",1,DT_F32_0731_FROM_OCC1,False)
Call TakeScreenShot
Call ClickBUtton("Process Open Items   \(Shift\+F4\)",False)
Call ClickBUtton("Select All",False)
Call ClickBUtton("Activate Items",False)
Call TakeScreenShot

Call ClickButton("Post   \(Ctrl\+S\)",False)
'Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot

Call GetStatusBar("item1","DT_F32_0131_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document "&DT_F32_0131_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" was posted in company code GR02")
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet("DT_F32_0131_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_F32_0131_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

'''''--------TransactionCode-FBL5N ----------''''

Call SetTcode(DT_F32_0131_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SelectRadioButton("X_CLSEL","Cleared items",False)
Call SetTextbox("Clearing date","SO_AUGDT-LOW",0,ConvertDate(DT_F32_1000_CLEARING_DATE),False)
Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_F32_1000_CUSTOMER_ACCOUNT,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)

Call SelectColumnGuiGrid("",0,"Document Number", False)
Call ClickButton("Set filter   \(Ctrl\+Shift\+F2\)",False)
Call TakeScreenShot

Call SetTextbox("Document Number","%%DYN001-LOW",0,DT_F32_1105_CLEARING_DOCUMENT,True)
'Call SetTextbox("Document Number","%%DYN001-LOW",0,DT_F32_0731_FROM,True)
'Call SetTextbox("to","%%DYN001-HIGH",0,DT_F32_1105_CLEARING_DOCUMENT,True)
Call TakeScreenShot

Call ClickBUtton("Execute   \(Enter\)",False)
Call TakeScreenShot

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_F32_0841_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",DT_F32_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_F32_0841_SEARCH_TERM_OCC1,True)
Call SetComboByKey("Search Direction",DT_F32_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter

Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_F32_0841_SEARCH_TERM_OCC2,True)
Call SetComboByKey("Search Direction",DT_F32_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter

Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonIfExist("Transfer   \(Enter\)",True)

DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_IP = year(DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)& ConvertDoubledigit(CSTR(Month(DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)))&ConvertDoubledigit(CSTR(Day(DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)))&DT_FORMULA_PCNTR
Call VerifyGridCellContent("", 1, "ZUONR", 0, DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_IP)
DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR_IP = year(DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR)& ConvertDoubledigit(CSTR(Month(DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR)))&ConvertDoubledigit(CSTR(Day(DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR)))&DT_FORMULA_PCNTR
Call VerifyGridCellContent("", 2, "ZUONR", 0, DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR_IP)
'DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_ZUONR_IP = year(DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_ZUONR)& ConvertDoubledigit(CSTR(Month(DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_ZUONR)))&ConvertDoubledigit(CSTR(Day(DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_ZUONR)))&DT_FORMULA_PCNTR
'Call VerifyGridCellContent("", 3, "ZUONR", 0, DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_ZUONR_IP)
'DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_ZUONR_IP = year(DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_ZUONR)& ConvertDoubledigit(CSTR(Month(DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_ZUONR)))&ConvertDoubledigit(CSTR(Day(DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_ZUONR)))&DT_FORMULA_PCNTR
'Call VerifyGridCellContent("", 4, "ZUONR", 0, DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_ZUONR_IP)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BELNR", 0, DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BELNR)
Call VerifyGridCellContent("", 2, "BELNR", 0, DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BELNR)
'Call VerifyGridCellContent("", 3, "BELNR", 0, DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BELNR)
'Call VerifyGridCellContent("", 4, "BELNR", 0, DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BELNR)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BLART", 0, DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BLART)
Call VerifyGridCellContent("", 2, "BLART", 0, DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BLART)
'Call VerifyGridCellContent("", 3, "BLART", 0, DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BLART)
'Call VerifyGridCellContent("", 4, "BLART", 0, DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BLART)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BLDAT", 0, ConvertDate(DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))
Call VerifyGridCellContent("", 2, "BLDAT", 0, ConvertDate(DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BLDAT))
'Call VerifyGridCellContent("", 3, "BLDAT", 0, ConvertDate(DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BLDAT))
'Call VerifyGridCellContent("", 4, "BLDAT", 0, ConvertDate(DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BLDAT))
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 2, "DMSHB", 0, DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB)
'Call VerifyGridCellContent("", 3, "DMSHB", 0, DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_DMSHB)
'Call VerifyGridCellContent("", 4, "DMSHB", 0, DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_DMSHB)
Call TakeScreenShot


Call VerifyGridCellContent("", 1, "HWAER", 0, DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER)
Call VerifyGridCellContent("", 2, "HWAER", 0, DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HWAER)
'Call VerifyGridCellContent("", 3, "HWAER", 0, DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_HWAER)
'Call VerifyGridCellContent("", 4, "HWAER", 0, DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_HWAER)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "AUGBL", 0, DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AUGBL)
Call VerifyGridCellContent("", 2, "AUGBL", 0, DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AUGBL)
'Call VerifyGridCellContent("", 3, "AUGBL", 0, DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AUGBL)
'Call VerifyGridCellContent("", 4, "AUGBL", 0, DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_AUGBL)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BUDAT", 0, ConvertDate(DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUDAT))
Call VerifyGridCellContent("", 2, "BUDAT", 0, ConvertDate(DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUDAT))
'Call VerifyGridCellContent("", 3, "BUDAT", 0, ConvertDate(DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUDAT))
'Call VerifyGridCellContent("", 4, "BUDAT", 0, ConvertDate(DT_F32_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUDAT))
Call TakeScreenShot


Call LogOff()
Call FinalStatus ()


