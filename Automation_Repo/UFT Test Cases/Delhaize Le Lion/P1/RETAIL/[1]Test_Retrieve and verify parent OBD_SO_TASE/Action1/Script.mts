'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_F-44-Clearing vendor 
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_RetrieveverifyOBD_SO_TASE"
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
'----------------TransactionCode-VA03-----------'
Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Order","VBAK-VBELN","",DT_VA03_0102_ORDER,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Display document flow   \(F5\)",False)
Call TakeScreenShot
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_SEARCH,True)
Call ClickButton("btn\[0\]",True)
Call TakeScreenShot
Call ClickButton("btn\[12\]",True)
Call ClickButton("Details   \(F2\)",False)
Call TakeScreenShot
Call SelectNodeGuiTree(0,"#1;#1")
Call TakeScreenShot
Call DoubleClick()
Call GetGridContent("Outb.*",0,"Doc.no.",1,"Item","","DT_VA03_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DOCNUM_Output")
' GetGridContent(gridTitle, gridIndex, columnName, rowNumber, refColumn, refFieldVal, dataTableColumnName)
' SelectRowGuiGrid(gridTitle, gridIndex, columnName, refValue, blnIsItPopup)
'Call SelectRowGuiGrid("Outb.*",0,"Item","10",False)
'Call TakeScreenShot
'Call ClickButtonToolBar("&DETAIL",0)
'Call TakeScreenShot
'Call SelectRowGuiGrid("",0,"Group description","",True)
'Call DoubleClick()
'Call ClickButton("Close window   \(Enter\)",True)
'SAPGuiSession("Session").SAPGuiWindow("Document Flow").Maximize
'SAPGuiSession("Session").SAPGuiWindow("Document Flow").SAPGuiToolbar("GridToolbar").PressButton "&DETAIL"
Call TakeScreenShot
Call ClickButton("Display document   \(F8\)",False)
Call TakeScreenShot
' VerifyTableCellContent(tableRowNumber, tableColumnName, tableName, expectedValue)
Call VerifyTableCellContent("1","Article","SAPMV50ATC_LIPS_OVER",DT_VA03_1102_CHECK_TEXT_OF_TABLECELL_ARTICLE_0)
Call VerifyTableCellContent("2","Article","SAPMV50ATC_LIPS_OVER",DT_VA03_1102_CHECK_TEXT_OF_TABLECELL_ARTICLE_1)
Call VerifyTableCellContent("1","Deliv. Qty","SAPMV50ATC_LIPS_OVER",DT_VA03_1102_CHECK_TEXT_OF_TABLECELL_DELIV_QTY_0)
Call VerifyTableCellContent("2","Deliv. Qty","SAPMV50ATC_LIPS_OVER",DT_VA03_1102_CHECK_TEXT_OF_TABLECELL_DELIV_QTY_1)
Call VerifyTableCellContent("1","Site","SAPMV50ATC_LIPS_OVER",DT_VA03_1102_CHECK_TEXT_OF_TABLECELL_SITE_0)
Call VerifyTableCellContent("2","Site","SAPMV50ATC_LIPS_OVER",DT_VA03_1102_CHECK_TEXT_OF_TABLECELL_SITE_1)
Call VerifyTableCellContent("1","SLoc","SAPMV50ATC_LIPS_OVER",DT_VA03_1102_CHECK_TEXT_OF_TABLECELL_SLOC_0)
Call VerifyTableCellContent("2","Sloc","SAPMV50ATC_LIPS_OVER",DT_VA03_1102_CHECK_TEXT_OF_TABLECELL_SLOC_1)
Call VerifyTextboxContent("Ship-To Party","KUWEV-KUNNR","",DT_VA03_1502_CHECK_TEXT_OF_SHIPTO_PARTY,False)
Call SelectTab("TAXI_TABSTRIP_OVERVIEW","Picking",False)
Call TakeScreenShot
Call GetTextboxValue("LIKP-BLDAT","","DT_VA03_1104_GET_PROPERTY_DOC_DATE",False)
Call TakeScreenShot
' GetTextboxValue(textboxName, textboxIndex, dataTableColumnName, blnIsItPopup)
Call VerifyTextBoxContent("Pick Date/Time","LIKP-KODAT","",DT_VA03_1104_CHECK_TEXT_OF_PICK_DATETIME,False)


''------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

