		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.06.01.04.02 Reverse AR Document

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

gstrTestCaseName = "Test_09.06.01.04.02 Reverse AR Document"
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

'''''''--------TransactionCode-FB08----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call TakeScreenShot

Call SelectCheckbox("TESTLAUF", "0", DT_F80_1000_TEST_RUN, False)
Call SetTextbox("Company code","BR_BUKRS-LOW","",DT_F80_1000_COMPANY_CODE,False)
Call SetTextbox("Reason for reversal","STOGRD","",DT_F80_1000_REASON_FOR_REVERSAL,False)
Call SetTextbox("Fiscal Year","BR_GJAHR-LOW","",DT_F80_1000_FISCAL_YEAR,False)
Call TakeScreenShot

Call ClickButton("%_BR_BELNR_%_APP_%-VALU_PUSH",False)
Call TakeScreenShot
Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_F80_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_F80_3010_TABLECELL_SINGLE_VALUE_1,True)
Call SetTableData("SAPLALDBSINGLE","Single value","3","","",DT_F80_3010_TABLECELL_SINGLE_VALUE_2,True)

Call ClickButtonIfExist("Copy   \(F8\)",True)

Call SetTextbox("Fiscal Year","BR_GJAHR-LOW","",DT_F80_1000_FISCAL_YEAR,False)
Call SetTextbox("Posting date","BR_BUDAT-LOW","",ConvertDate(DT_F80_1000_POSTING_DATE),False)
Call SetTextbox("to","BR_BUDAT-HIGH","",ConvertDate(DT_F80_1000_TO),False)
Call SetTextbox("Posting Date","STODAT","",ConvertDate(DT_F80_1000_POSTING_DATE_OCC1),False)
Call SetTextbox("Posting period","MONAT","",ConvertDoubledigit(DT_F80_1000_POSTING_PERIOD),False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call VerifyifGuiLabelExistsByRelativeid(DT_F80_0120_CHECK_TEXT_OF_NO_NAME, "wnd\[0\]/usr/lbl\[1,6\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_F80_0120_CHECK_TEXT_OF_NO_NAME_OCC1, "wnd\[0\]/usr/lbl\[1,7\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_F80_0120_CHECK_TEXT_OF_NO_NAME_OCC2, "wnd\[0\]/usr/lbl\[1,8\]")

Call VerifyifGuiLabelExistsByRelativeid(ConvertDate(DT_F80_0120_CHECK_TEXT_OF_NO_NAME_OCC3), "wnd\[0\]/usr/lbl\[81,6\]")
Call VerifyifGuiLabelExistsByRelativeid(ConvertDate(DT_F80_0120_CHECK_TEXT_OF_NO_NAME_OCC4), "wnd\[0\]/usr/lbl\[81,7\]")
Call VerifyifGuiLabelExistsByRelativeid(ConvertDate(DT_F80_0120_CHECK_TEXT_OF_NO_NAME_OCC5), "wnd\[0\]/usr/lbl\[81,8\]")

Call VerifyifGuiLabelExistsByRelativeid(ConvertDate(DT_F80_0120_CHECK_TEXT_OF_NO_NAME_OCC6), "wnd\[0\]/usr/lbl\[65,6\]")
Call VerifyifGuiLabelExistsByRelativeid(ConvertDate(DT_F80_0120_CHECK_TEXT_OF_NO_NAME_OCC7), "wnd\[0\]/usr/lbl\[65,7\]")
Call VerifyifGuiLabelExistsByRelativeid(ConvertDate(DT_F80_0120_CHECK_TEXT_OF_NO_NAME_OCC8), "wnd\[0\]/usr/lbl\[65,8\]")
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)
Call SelectCheckbox("TESTLAUF", "0", DT_F80_1000_TEST_RUN_OCC1, False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call GetLabelContentByRefLabel("Document Number", -112, -64, "DT_F80_0120_CHECK_TEXT_OF_REVERSED_WITH_DOCUMENT_3400000018_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_F80_0120_CHECK_TEXT_OF_REVERSED_WITH_DOCUMENT_3400000018_OUTPUT",DT_F80_0120_CHECK_TEXT_OF_REVERSED_WITH_DOCUMENT_3400000018)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call GetLabelContentByRefLabel("Document Number", -112, -96, "DT_F80_0120_CHECK_TEXT_OF_REVERSED_WITH_DOCUMENT_3400000019_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_F80_0120_CHECK_TEXT_OF_REVERSED_WITH_DOCUMENT_3400000019_OUTPUT",DT_F80_0120_CHECK_TEXT_OF_REVERSED_WITH_DOCUMENT_3400000019)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call GetLabelContentByRefLabel("Document Number", -112, -128, "DT_F80_0120_CHECK_TEXT_OF_REVERSED_WITH_DOCUMENT_3400000020_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_F80_0120_CHECK_TEXT_OF_REVERSED_WITH_DOCUMENT_3400000020_OUTPUT",DT_F80_0120_CHECK_TEXT_OF_REVERSED_WITH_DOCUMENT_3400000020)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call VerifyifGuiLabelExistsByRelativeid(DT_F80_0120_CHECK_TEXT_OF_NO_NAME_OCC9, "wnd\[0\]/usr/lbl\[1,7\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_F80_0120_CHECK_TEXT_OF_NO_NAME_OCC10, "wnd\[0\]/usr/lbl\[1,9\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_F80_0120_CHECK_TEXT_OF_NO_NAME_OCC11, "wnd\[0\]/usr/lbl\[1,11\]")

Call VerifyifGuiLabelExistsByRelativeid(Lcase(DT_F80_0120_CHECK_TEXT_OF_REVERSED_WITH_DOCUMENT_3400000018_OCC1), "wnd\[0\]/usr/lbl\[17,8\]")
Call VerifyifGuiLabelExistsByRelativeid(Lcase(DT_F80_0120_CHECK_TEXT_OF_REVERSED_WITH_DOCUMENT_3400000019_OCC1), "wnd\[0\]/usr/lbl\[17,10\]")
Call VerifyifGuiLabelExistsByRelativeid(Lcase(DT_F80_0120_CHECK_TEXT_OF_REVERSED_WITH_DOCUMENT_3400000020_OCC1), "wnd\[0\]/usr/lbl\[17,12\]")
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
'''''''''--------TransactionCode-FBL5N ----------''''

Call SetTcode(DT_F80_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_F80_1000_CUSTOMER_ACCOUNT,False)
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_F80_1000_COMPANY_CODE_OCC1,False)
Call SelectRadioButton("X_AISEL", "All items", False)
Call TakeScreenShot

Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call ClickButton("%_%%DYN011_%_APP_%-VALU_PUSH",False)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_F80_3010_TABLECELL_SINGLE_VALUE_0_OCC1,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_F80_3010_TABLECELL_SINGLE_VALUE_1_OCC1,True)
Call SetTableData("SAPLALDBSINGLE","Single value","3","","",DT_F80_3010_TABLECELL_SINGLE_VALUE_2_OCC1,True)
Call SetTableData("SAPLALDBSINGLE","Single value","4","","",DT_F80_3010_TABLECELL_SINGLE_VALUE_3,True)
Call SetTableData("SAPLALDBSINGLE","Single value","5","","",DT_F80_3010_TABLECELL_SINGLE_VALUE_4,True)
Call SetTableData("SAPLALDBSINGLE","Single value","6","","",DT_F80_3010_TABLECELL_SINGLE_VALUE_5,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_F80_0841_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",DT_F80_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_F80_0841_SEARCH_TERM_OCC1,True)
Call SetComboByKey("Search Direction",DT_F80_0841_SEARCH_DIRCT_OCC1)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_F80_0841_SEARCH_TERM_OCC2,True)
Call SetComboByKey("Search Direction",DT_F80_0841_SEARCH_DIRCT_OCC2)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_F80_0841_SEARCH_TERM_OCC3,True)
Call SetComboByKey("Search Direction",DT_F80_0841_SEARCH_DIRCT_OCC3)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)


Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_F80_0841_SEARCH_TERM_OCC4,True)
Call SetComboByKey("Search Direction",DT_F80_0841_SEARCH_DIRCT_OCC4)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_F80_0841_SEARCH_TERM_OCC5,True)
Call SetComboByKey("Search Direction",DT_F80_0841_SEARCH_DIRCT_OCC5)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonIfExist("Transfer   \(Enter\)",True)

Call VerifyGridCellContent("", 1, "ZUONR", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)
Call VerifyGridCellContent("", 2, "ZUONR", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR)
Call VerifyGridCellContent("", 3, "ZUONR", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_ZUONR)
Call VerifyGridCellContent("", 4, "ZUONR", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_ZUONR)
Call VerifyGridCellContent("", 5, "ZUONR", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_ZUONR)
Call VerifyGridCellContent("", 6, "ZUONR", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_5_ZUONR)

Call VerifyGridCellContent("", 1, "BELNR", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 2, "BELNR", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BELNR)
Call VerifyGridCellContent("", 3, "BELNR", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BELNR)
Call VerifyGridCellContent("", 4, "BELNR", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BELNR)
Call VerifyGridCellContent("", 5, "BELNR", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BELNR)
Call VerifyGridCellContent("", 6, "BELNR", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_5_BELNR)

Call VerifyGridCellContent("", 1, "BLART", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 2, "BLART", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BLART)
Call VerifyGridCellContent("", 3, "BLART", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BLART)
Call VerifyGridCellContent("", 4, "BLART", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BLART)
Call VerifyGridCellContent("", 5, "BLART", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BLART)
Call VerifyGridCellContent("", 6, "BLART", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_5_BLART)

Call VerifyGridCellContent("", 1, "BLDAT", 0, ConvertDate(DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))
Call VerifyGridCellContent("", 2, "BLDAT", 0, ConvertDate(DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BLDAT))
Call VerifyGridCellContent("", 3, "BLDAT", 0, ConvertDate(DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BLDAT))
Call VerifyGridCellContent("", 4, "BLDAT", 0, ConvertDate(DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BLDAT))
Call VerifyGridCellContent("", 5, "BLDAT", 0, ConvertDate(DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BLDAT))
Call VerifyGridCellContent("", 6, "BLDAT", 0, ConvertDate(DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_5_BLDAT))

Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 2, "DMSHB", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB)
Call VerifyGridCellContent("", 3, "DMSHB", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_DMSHB)
Call VerifyGridCellContent("", 4, "DMSHB", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_DMSHB)
Call VerifyGridCellContent("", 5, "DMSHB", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_DMSHB)
Call VerifyGridCellContent("", 6, "DMSHB", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_5_DMSHB)

Call VerifyGridCellContent("", 1, "AUGBL", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AUGBL)
Call VerifyGridCellContent("", 2, "AUGBL", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AUGBL)
Call VerifyGridCellContent("", 3, "AUGBL", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AUGBL)
Call VerifyGridCellContent("", 4, "AUGBL", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_AUGBL)
Call VerifyGridCellContent("", 5, "AUGBL", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_AUGBL)
Call VerifyGridCellContent("", 6, "AUGBL", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_5_AUGBL)

Call VerifyGridCellContent("", 1, "ZTERM", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZTERM)
Call VerifyGridCellContent("", 2, "ZTERM", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZTERM)
Call VerifyGridCellContent("", 3, "ZTERM", 0, "")
Call VerifyGridCellContent("", 4, "ZTERM", 0, "")
Call VerifyGridCellContent("", 5, "ZTERM", 0, "")
Call VerifyGridCellContent("", 6, "ZTERM", 0, "")

Call VerifyGridCellContent("", 1, "BUDAT", 0, ConvertDate(DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUDAT))
Call VerifyGridCellContent("", 2, "BUDAT", 0, ConvertDate(DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUDAT))
Call VerifyGridCellContent("", 3, "BUDAT", 0, ConvertDate(DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BUDAT))
Call VerifyGridCellContent("", 4, "BUDAT", 0, ConvertDate(DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BUDAT))
Call VerifyGridCellContent("", 5, "BUDAT", 0, ConvertDate(DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BUDAT))
Call VerifyGridCellContent("", 6, "BUDAT", 0, ConvertDate(DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_5_BUDAT))

Call VerifyGridCellContent("", 1, "PRCTR", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("", 2, "PRCTR", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)
Call VerifyGridCellContent("", 3, "PRCTR", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_PRCTR)

Call VerifyGridCellContent("", 1, "HKONT", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
Call VerifyGridCellContent("", 2, "HKONT", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)
Call VerifyGridCellContent("", 3, "HKONT", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_HKONT)
Call VerifyGridCellContent("", 4, "HKONT", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_HKONT)
Call VerifyGridCellContent("", 5, "HKONT", 0, DT_HKONT_4)
Call VerifyGridCellContent("", 6, "HKONT", 0, DT_HKONT_5)

Call LogOff()
Call FinalStatus()



