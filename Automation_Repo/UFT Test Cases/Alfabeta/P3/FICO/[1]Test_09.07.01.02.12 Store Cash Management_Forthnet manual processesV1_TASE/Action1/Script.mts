		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.07.01.02.12 Store Cash Management_Forthnet manual processesV1
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

gstrTestCaseName = "Test_09.07.01.02.12 Store Cash Management_Forthnet manual processesV1"
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
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F21_0100_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F21_0100_DOCHEADER_TEXT,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F21_0100_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F21_0301_AMOUNT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F21_0301_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F21_0301_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F21_0301_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F21_0300_AMOUNT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F21_0300_TEXT,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_F21_1006_BUSINESS_AREA,False)
Call SetTextbox("Profit Center","COBL-PRCTR","",DT_F21_1006_PROFIT_CENTER,False)

Call SelectMenuBar("Document;Simulate")

Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)

Call GetStatusBar("item1","DT_F21_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document "&DT_F21_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" was posted in company code GR02")
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet("DT_F21_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_F21_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)


''''''--------TransactionCode-FBL5N----------''''

Call SetTcode(DT_F21_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot


Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_F21_1000_CUSTOMER_ACCOUNT,False)
Call SelectRadioButton("X_AISEL","All items", False)
Call TakeScreenShot

Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call ClickButton("%_%%DYN011_%_APP_%-VALU_PUSH",False)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_F21_1106_DOCUMENT_NUMBER,True)
Call TakeScreenShot
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)

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

Call VerifyGridCellContent("", 1, "BELNR", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 1, "BLART", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 1, "BLDAT", 0, ConvertDate(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))
Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 1, "SGTXT", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
Call VerifyGridCellContent("", 1, "BUDAT", 0, ConvertDate(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUDAT))
Call VerifyGridCellContent("", 1, "HKONT", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
Call VerifyGridCellContent("", 1, "PRCTR", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call LogOff'
