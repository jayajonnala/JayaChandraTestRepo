'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_02-06-01-18-QMatics excel upload 
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_02-06-01-18-QMatics excel"
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
' StartExecution(excelPath, strTestCaseName, iterationIndex, strResultFolderPath)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''''''--------TransactionCode-ZMDPC_UPLOAD_COND----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call ClickButton("Get Variant...   \(Shift\+F5\)",False)
Call TakeScreenShot
Call SetTextbox("Variant","V-LOW","",DT_ZMDPC_UPLOAD_COND_0841_SEARCH_TERM,True)
Call SetTextbox("Created By","ENAME-LOW",0,"",True)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",True)
Call TakeScreenShot

Call SetTextbox("File name","P_FILE","",DT_ZMDPC_UPLOAD_COND_1000_FILE_NAME,False)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "Article", "", DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR)
Call VerifyGridCellContent("", 1, "Display Unit/Measure", "", DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MEINH)
Call VerifyGridCellContent("", 1, "Supplier", "", DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LIFNR)
Call VerifyGridCellContent("", 1, "File Price", "", DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KBETR_FILE)
Call VerifyGridCellContent("", 1, "Price(Editable)", "", DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KBETR_EDIT)
Call VerifyGridCellContent("", 1, "Condition currency", "", DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KONWA)
Call VerifyGridCellContent("", 1, "Pricing unit", "", DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KPEIN)
Call TakeScreenShot

Call SelectRowGuiGridbyRowNo("", "",1, False)
Call TakeScreenShot
Call ClickButton("Create Conditions   \(F8\)",False)
Call TakeScreenShot

Call VerifyStatusBar(DT_ZMDPC_UPLOAD_COND_0500_CHECK_TEXT_OF_STATUSBAR)

''''''--------TransactionCode-nVKP5----------''''

Call SetTcode(DT_ZMDPC_UPLOAD_COND_0500_OKCD)     
Call PressEnter() 
Call TakeScreenShot


Call SetTextbox("Purchase Price Determ. Seq.","P_EKERV","",DT_ZMDPC_UPLOAD_COND_1000_PURCHASE_PRICE_DETERM_SEQ,False)
Call SetTextbox("List Group","P_LIGRU","",DT_ZMDPC_UPLOAD_COND_1000_LIST_GROUP,False)
Call SetTextbox("Article","S_MATNR-LOW","",DT_ZMDPC_UPLOAD_COND_1000_ARTICLE,False)
Call SetTextbox("Sales organization","S_VKORG-LOW","",DT_ZMDPC_UPLOAD_COND_1000_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution channel","S_VTWEG-LOW","",DT_ZMDPC_UPLOAD_COND_1000_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Validity","S_DATUM-LOW","",ConvertDate(DT_ZMDPC_UPLOAD_COND_1000_VALIDITY),False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call PressEnter()
Call VerifyGridCellContent("", 1, "Supplier", "", DT_ZMDPC_UPLOAD_COND_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LIFNR)
 Call VerifyGridCellContent("", 1, "Article", "", DT_ZMDPC_UPLOAD_COND_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR)
 Call VerifyGridCellContent("", 1, "Sales Organization", "", DT_ZMDPC_UPLOAD_COND_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VKORG)
 Call VerifyGridCellContent("", 1, "Distribution Channel", "", DT_ZMDPC_UPLOAD_COND_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VTWEG)
 Call VerifyGridCellContent("", 1, "Base Unit of Measure", "", DT_ZMDPC_UPLOAD_COND_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MEINS)
 Call VerifyGridCellContent("", 1, "Basic Pur.Pr.", "", DT_ZMDPC_UPLOAD_COND_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_EKPGR)
 Call VerifyGridCellContent("", 1, "Purch.net/net", "", DT_ZMDPC_UPLOAD_COND_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_EKPNN)
Call VerifyGridCellContent("", 1, "Final price", "", DT_ZMDPC_UPLOAD_COND_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ENDPA)
Call VerifyGridCellContent("", 1, "Sales Pr. from Cond.", "", DT_ZMDPC_UPLOAD_COND_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ENDPA)

' ClickButtonToolBar(buttonName, toolbarIndex)
Call ClickButtonToolBar ("EKDT",0)
Call VerifyTextBoxContent("Material","KOMP-MATNR","",DT_ZMDPC_UPLOAD_COND_9000_CHECK_TEXT_OF_ARTICLE ,False)
''' VerifyTableCellContent(tableRowNumber, tableColumnName, tableName, expectedValue)
Call VerifyTableCellContent(9,"CnTy", "SAPLV69ATCTRL_KONDITIONEN", DT_ZMDPC_UPLOAD_COND_6201_CHECK_TEXT_OF_TABLECELL_CNTY_8)
Call VerifyTableCellContent(9,"Description", "SAPLV69ATCTRL_KONDITIONEN", DT_ZMDPC_UPLOAD_COND_6201_CHECK_TEXT_OF_TABLECELL_NAME_8)
Call VerifyTableCellContent(9,"Amount", "SAPLV69ATCTRL_KONDITIONEN", DT_ZMDPC_UPLOAD_COND_6201_CHECK_TEXT_OF_TABLECELL_AMOUNT_8)
Call VerifyTableCellContent(9,"Crcy", "SAPLV69ATCTRL_KONDITIONEN", DT_ZMDPC_UPLOAD_COND_6201_CHECK_TEXT_OF_TABLECELL_CRCY_8)
Call VerifyTableCellContent(9,"per", "SAPLV69ATCTRL_KONDITIONEN", DT_ZMDPC_UPLOAD_COND_6201_CHECK_TEXT_OF_TABLECELL_PER_8)
Call VerifyTableCellContent(9,"Condition value", "SAPLV69ATCTRL_KONDITIONEN", DT_ZMDPC_UPLOAD_COND_6201_CHECK_TEXT_OF_TABLECELL_CONDITION_VALUE_8)
Call VerifyTableCellContent(9,"Non-Active", "SAPLV69ATCTRL_KONDITIONEN", "S_LEDG")
Call ClickButton("Back   \(F3\)",False)
Call ClickButtonIfExist("No", True)
Call ClickButton("Save   (Ctrl+S)",False)

Call GetStatusBar("item1", "DT_VKP5_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call VerifyStatusBar(DT_ZMDPC_UPLOAD_COND_1000_CHECK_TEXT_OF_STATUSBAR)

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()




'
'
'
'
'
'
'SAPGuiSession("Session").SAPGuiWindow("Create Purchasing: Item").SAPGuiTable("Pricing Elements").SetCellData
'
'Call ClickContextButtonToolBar("&MB_VARIANT",1)
'Call SelectMenuIdToolBar("&COL0",1)
'Call ClickButtonToolBar("&FIND",0)
'
'''Call ClickButtonToolBar("&FIND", 1)
'Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Previous Tot Marg (%)",True)
'Call TakeScreenShot
'Call SetCombo("GS_SEARCH-SEARCH_ORDER","From Start of Table Downwards")
'Call ClickButton("OK   \(Enter\)",True)
'Call ClickButton("Cancel   \(F12\)",True)
'Call ClickButton("Show Selected Fields \(F7\)",True)
'Call TakeScreenShot
'Call ClickButton("Transfer   (Enter)",True)
'
''Call GetGridContent("", "", "Total Margin (%)", 1, "", "", "DT_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MARGIN_DLL_T")
'Call TakeScreenShot
'Call ClickButton("Back   \(F3\)",False)
'Call ClickButtonIfExist("No", True)
''Call SetTextbox("Valid On Date","P_VKKAB","",ConvertDate(DT_ZMDPC_MARGIN_1000_VALID_ON_DATE_OCC1),False)
'Call TakeScreenShot
'Call ClickButton("Execute   \(F8\)",False)
'Call TakeScreenShot
''
''Call ClickButton("Change Layout...   \(Ctrl\+F8\)",False)
''Call ClickButtonToolBar("&FIND", 1)
'Call ClickContextButtonToolBar("&MB_VARIANT",1)
'Call SelectMenuIdToolBar("&COL0",1)
'Call ClickButtonToolBar("&FIND",0)
'
'Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","WS Price (ZPR0) PG",True)
'Call TakeScreenShot
'Call SetCombo("GS_SEARCH-SEARCH_ORDER","From Start of Table Downwards")
'Call ClickButton("Cancel   \(F12\)",True)
'Call ClickButton("Show Selected Fields \(F7\)",True)
'Call TakeScreenShot
'Call ClickButton("Transfer   (Enter)",True)
'
'
'Call VerifyGridCellContent("", 1, "Sales Org", "", DT_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VKORG_OCC1)
'Call VerifyGridCellContent("", 1, "Dist Chl", "", DT_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VTWEG_OCC1)
'Call VerifyGridCellContent("", 1, "PL", "", DT_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PLTYP_OCC1)
'Call VerifyGridCellContent("", 1, "Article", "", DT_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR_OCC1)
'Call TakeScreenShot
'
''''''Call VerifyGridCellContent("shell", 1, "Sales Org", "", DT_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VKORG_OCC1)
''''''Call VerifyGridCellContent("shell", 1, "Dist Chl", "", DT_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VTWEG_OCC1)
''''''Call VerifyGridCellContent("shell", 1, "PL", "", DT_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PLTYP_OCC1)
''''''Call VerifyGridCellContent("shell", 1, "Article", "", DT_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR_OCC1)
''''''Call TakeScreenShot
'
'Call GetGridContent("", "", "Net Value (ZPAF)", 1, "", "","DT_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_NET_PRICE")
'Call GetGridContent("", "", "WS Price (ZPR0) PG", 1, "", "","DT_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WHS_PRICE")
'Call GetGridContent("", "", "Efective Price(ZDPR)", 1, "", "","DT_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_EFE_PRICE")
'Call GetGridContent("", "", "Purch Price(PB00)", 1, "", "","DT_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BAS_PRICE")
'Call TakeScreenShot
'Call VerifyGridCellContent("", 1, "Retail Sales UOM", "","DT_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VRKME")
'Call VerifyGridCellContent("", 1, "Total Margin (%)", "","DT_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MARGIN_DLL_T_OCC1")

'*********************************************End Of Script*********************************************************************

