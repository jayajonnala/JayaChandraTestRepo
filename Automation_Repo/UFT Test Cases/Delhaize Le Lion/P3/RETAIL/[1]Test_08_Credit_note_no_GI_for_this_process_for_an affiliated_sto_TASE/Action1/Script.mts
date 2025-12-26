'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_08_Credit_note_no_GI_for_this_process_for_an affiliated_sto  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_08_Credit_note_affiliated"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\DLL_P3\Data\TASE_DT_02-04-01-05-03-Create new assortment-BASA.xls"

If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	datatable_row= Parameter("datatable_row")	
End If
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''--------TransactionCode-va03----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE) 

Call SetTextbox("Order","VBAK-VBELN","",DT_VA03_0102_ORDER,False)
Call PressEnter()     
Call TakeScreenShot
Call ClickButton("Display document flow   \(F5\)",False)
Call ActivateItemGuiTree(0, "#1;#1;#1", "#1")
Call GetGridContent("Accounting document.*", 0, "Doc.no.", 1, "Status", "Not Cleared", "DT_OP_VA03_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DOCNUM")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call ClickButton("Back   \(F3\)", False)
Call ClickButton("Back   \(F3\)", False)
Call ClickButton("Back   \(F3\)", False)

'''''''--------TransactionCode-/FB03----------''''


Call SetTcode(DT_VA03_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)


Call SetTextbox("Document Number","RF05L-BELNR","",DT_VA03_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_VA03_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_VA03_0100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()
Call ClickButtonToolBar("&MB_FILTER",0)
Call TakeScreenShot
'Call ClickButtonToolBar("&FIND",True)
Call SelectRowGuiGridbyRowNo("Column Set", 0, 34, True)
'Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Branch account",True)
'Call SetCombo("Search Direction","From Start of Table Downwards")
Call TakeScreenShot
'Call SelectCheckbox("GS_SEARCH-EXACT_WORD",1,"ON",True)
'Call ClickButtonIfExist("OK   \(Enter\)",True)
'Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Add Filter Criterion \(F7\)",True)
Call TakeScreenShot
'Call ClickButtonToolBar("&FIND",True)
Call SelectRowGuiGridbyRowNo("Column Set", 0, 42, True)
'Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","G/L Account",True)
''Call SetCombo("Search Direction","From Start of Table Downwards")
'Call TakeScreenShot
'Call SelectCheckbox("GS_SEARCH-EXACT_WORD",1,"ON",True)
'Call ClickButtonIfExist("OK   \(Enter\)",True)
'Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Add Filter Criterion \(F7\)",True)
'Call TakeScreenShot
'Call ClickButtonToolBar("&FIND",True)
'Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Account",True)
'Call SetCombo("Search Direction","From Start of Table Downwards")
Call ClickButtonIfExist("Transfer   \(Enter\)",True)
Call TakeScreenShot
Call ClickButtonToolBar("&MB_FILTER",0)
Call SelectRowGuiGridbyRowNo("Column Set", 0,8, True)
Call ClickButtonIfExist("Add Filter Criterion \(F7\)",True)
'Call SelectCheckbox("GS_SEARCH-EXACT_WORD",1,"ON",True)
'Call ClickButtonIfExist("OK   \(Enter\)",True)
'Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call TakeScreenShot
Call ClickButtonIfExist("Define Filter Values",True)
Call TakeScreenShot
Call ClickButtonIfExist("%_%%DYN001_%_APP_%-VALU_PUSH",True)
Call SetTableData("SAPLALDBSINGLE", "Single value",1, "","",DT_FB03_1105_ACCOUNT_MAINTENANCE,false)
Call SetTableData("SAPLALDBSINGLE", "Single value",2, "","",DT_FB03_3010_TABLECELL_SINGLE_VALUE_1_MAINTENANCE,false)
Call SetTableData("SAPLALDBSINGLE", "Single value",3, "","",DT_FB03_3010_TABLECELL_SINGLE_VALUE_2_MAINTENANCE, false)
Call TakeScreenShot
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(Enter\)",True)
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "Posting Key","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "Posting Key","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 3, "Posting Key","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("", 1, "Assignment","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)
'Call VerifyGridCellContent("", 2, "Assignment","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR)
'Call VerifyGridCellContent("", 3, "Assignment","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_ZUONR)
Call VerifyGridCellContent("", 1, "Tax Code","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ)
Call VerifyGridCellContent("", 2, "Tax Code","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MWSKZ)
Call VerifyGridCellContent("", 3, "Tax Code","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_MWSKZ)
Call VerifyGridCellContent("", 1, "Profit Center","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOSTL)
Call VerifyGridCellContent("", 2, "Profit Center","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KOSTL)
Call VerifyGridCellContent("", 3, "Profit Center","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KOSTL)
'Call VerifyGridCellContent("", 1, "Cost Center","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
'Call VerifyGridCellContent("", 2, "Cost Center","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)
'Call VerifyGridCellContent("", 3, "Cost Center","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_PRCTR)
Call ClickButtonToolBar("&MB_FILTER",0)
Call TakeScreenShot
Call ClickButtonIfExist("Remove Filter Criterion \(F6\)",True)
'Call ClickButtonToolBar("&FIND",TRUE)
Call SelectRowGuiGridbyRowNo("Column Set", 0,21, True)
'Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Trading Partner",True)
'Call SetCombo("Search Direction","From Start of Table Downwards")
'Call TakeScreenShot
'Call SelectCheckbox("GS_SEARCH-EXACT_WORD",1,"ON",True)
'Call ClickButtonIfExist("OK   \(Enter\)",True)
'Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Add Filter Criterion \(F7\)",True)
Call TakeScreenShot
Call ClickButtonIfExist("Transfer   \(Enter\)",True)
Call VerifyGridCellContent("", 1, "Trading partner","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VBUND)
Call VerifyGridCellContent("", 2, "Trading partner","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_VBUND)
Call VerifyGridCellContent("", 3, "Trading partner","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_VBUND)
Call ClickButtonIfExist("Display Document Header   \(F5\)",False)
Call VerifyTextBoxContent("Document type", "BKPF-BLART", "", DT_VA03_1710_CHECK_TEXT_OF_DOCUMENT_TYPE, True)
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

