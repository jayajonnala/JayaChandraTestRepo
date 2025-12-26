		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.06.01.01.07 Manage Manual Customer Credit Note

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

gstrTestCaseName = "Test_09.06.01.01.07 Manage Manual Customer Credit Note"
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

''''--------TransactionCode-FB75----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB75_1000_COMPANY_CODE,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SelectMenuBar("Goto;Post with Reference")
Call TakeScreenShot

Call SelectCheckbox("RF05A-CPBET","0",DT_FB75_0104_DO_NOT_PROPOSE_AMOUNTS,False)
Call SelectCheckbox("RF05A-CPTEX","0",DT_FB75_0104_COPY_TEXTS,False)

Call SetTextbox("Document Number","BKPF-BELNR", "", DT_FB75_0104_DOCUMENT_NUMBER, False)
Call SetTextbox("Company Code","BKPF-BUKRS", "", DT_FB75_0104_COMPANY_CODE, False)
Call SetTextbox("Fiscal Year","BKPF-GJAHR", "", DT_FB75_0104_FISCAL_YEAR, False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call SetTextbox("Reference","BKPF-XBLNR", "", "", False)
'Call SetTextbox("Reference","BKPF-XBLNR", "", DT_FB75_0100_REFERENCE, False)
Call PressEnter()
Call TakeScreenShot
Call SetTextbox("G/L Acc","BSEG-HKONT", "", DT_FB75_0301_GL_ACC, False)
Call SetTextbox("Amount","BSEG-WRBTR", "", DT_FB75_0301_AMOUNT, False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call SetTextbox("Amount","BSEG-WRBTR", "", DT_FB75_0300_AMOUNT, False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call SelectMenuBar("Document;Simulate")
Call TakeScreenShot
Call PressEnter()
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)

Call GetTextboxValue("MESSTXT1", "0", "DT_FB75_0010_CHECK_TEXT_OF_MESSTXT1_OUTPUT", True)
Call WriteRunTimeDataToExcelGlobalSheet("DT_FB75_0010_CHECK_TEXT_OF_MESSTXT1_OUTPUT",DT_FB75_0010_CHECK_TEXT_OF_MESSTXT1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyTextBoxContent("Information Message","MESSTXT1", "0",Lcase(DT_FB75_0010_CHECK_TEXT_OF_MESSTXT1_OCC1), True)

Call GetTextboxValue("MESSTXT2", "0", "DT_FB75_0010_CHECK_TEXT_OF_MESSTXT2_OUTPUT", True)
Call WriteRunTimeDataToExcelGlobalSheet("DT_FB75_0010_CHECK_TEXT_OF_MESSTXT2_OUTPUT",DT_FB75_0010_CHECK_TEXT_OF_MESSTXT2)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyTextBoxContent("Information Message","MESSTXT2", "0", Lcase(DT_FB75_0010_CHECK_TEXT_MESSTXT2_OCC1), True)

Call ClickButton("Continue   \(Enter\)",True)
Call ClickButtonIfExist("Cancel   \(F12\)",False)

'''''''''--------TransactionCode-FBL5N ----------''''

Call SetTcode(DT_FB75_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_FB75_1000_CUSTOMER_ACCOUNT,False)
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_FB75_1000_COMPANY_CODE_OCC1,False)
Call TakeScreenShot

Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call ClickButton("%_%%DYN011_%_APP_%-VALU_PUSH",False)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_FB75_1106_DOCUMENT_NUMBER,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB75_0841_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",DT_FB75_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB75_0841_SEARCH_TERM_OCC1,True)
Call SetComboByKey("Search Direction",DT_FB75_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB75_0841_SEARCH_TERM_OCC2,True)
Call SetComboByKey("Search Direction",DT_FB75_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB75_0841_SEARCH_TERM_OCC3,True)
Call SetComboByKey("Search Direction",DT_FB75_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB75_0841_SEARCH_TERM_OCC4,True)
Call SetComboByKey("Search Direction",DT_FB75_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB75_0841_SEARCH_TERM_OCC5,True)
Call SetComboByKey("Search Direction",DT_FB75_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB75_0841_SEARCH_TERM_OCC6,True)
Call SetComboByKey("Search Direction",DT_FB75_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonIfExist("Transfer   \(Enter\)",True)

Call VerifyGridCellContent("", 1, "ZUONR", 0, DT_FB75_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)
Call VerifyGridCellContent("", 1, "BELNR", 0, DT_FB75_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 1, "BLART", 0, DT_FB75_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 1, "BLDAT", 0, ConvertDate(DT_FB75_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))
Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_FB75_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 1, "SGTXT", 0, DT_FB75_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
Call VerifyGridCellContent("", 1, "ZTERM", 0, "")
Call VerifyGridCellContent("", 1, "KOSTL", 0, "")
Call VerifyGridCellContent("", 1, "BUDAT", 0, ConvertDate(DT_FB75_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUDAT))
Call VerifyGridCellContent("", 1, "PRCTR", 0, DT_FB75_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("", 1, "HKONT", 0, DT_FB75_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)

Call TakeScreenShot

''''--------TransactionCode-F.62 ----------''''

Call SetTcode(DT_FB75_0500_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SelectCheckbox("NORMBL", 0, DT_FB75_1000_STANDARD_DOCUMENTS, False)
Call SetTextbox("Fiscal year","RGJAHR-LOW","",DT_FB75_1000_FISCAL_YEAR,False)
Call SetTextbox("Company code","RBUKRS-LOW","",DT_FB75_1000_COMPANY_CODE_OCC2,False)
Call SetTextbox("Document number","RBELNR-LOW","",DT_FB75_1000_DOCUMENT_NUMBER,False)
Call SetTextbox("Document type","RBLART-LOW","",DT_FB75_1000_DOCUMENT_TYPE,False)
Call SetTextbox("Correspondence","REVENT","",DT_FB75_1000_CORRESPONDENCE,False)
Call SetTextbox("Posting date","RBUDAT-LOW","",ConvertDate(DT_FB75_1000_POSTING_DATE),False)
Call SetTextbox("Document date","RBLDAT-LOW","",ConvertDate(DT_FB75_1000_DOCUMENT_DATE),False)
Call SetTextbox("Delete if finished since","RDELDAYS","","",False)
'Call SetTextbox("Delete if finished since","RDELDAYS","",DT_FB75_1000_DELETE_IF_FINISHED_SINCE,False)
Call SetTextbox("Log to printer","PRDEST","",DT_FB75_1000_LOG_TO_PRINTER,False)
Call ClickBUtton("Execute   \(F8\)",False)
Call VerifyTextBoxContent("Information Message","MESSTXT1", "0",Lcase(DT_FB75_0010_CHECK_TEXT_OF_MESSTXT1_OCC2), True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call ClickButtonIfExist("Yes",True)
'Call SetTextbox("Output to Printer","TDDEST","",DT_FB75_1000_LOG_TO_PRINTER,False)
Call PressEnter()
Call TakeScreenShot

''''--------TransactionCode-SP02 ----------''''

Call SetTcode(DT_FB75_0120_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetFocusGuiLabel(DT_FB75_0120_CHECK_TEXT_OF_NO_NAME,"347","56", False)
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenshot
Call VerifyifGuiLabelExistsByRelativeid(DT_FB75_0120_CHECK_TEXT_OF_NO_NAME,"wnd\[0\]/usr/lbl\[49,3\]")
Call TakeScreenshot
Call ClickLabel("X__PDF", "0", False)
'Wait time added for the loading of PDF screen
Wait(10)
Call TakeScreenshot

Call LogOff()
Call FinalStatus()




