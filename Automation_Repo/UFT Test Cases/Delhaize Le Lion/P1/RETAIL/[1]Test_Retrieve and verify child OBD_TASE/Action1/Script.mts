

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_159_Update Account Completion Table (All Opcos)_TASE
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
	DataRowSet= Parameter("datatable_row")	
End If
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_Retrieve and verify child OBD"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''' Login '''
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''SAP Login'''
'
Call SetTcode(DT_SAPTRANSACTIONCODE)     ' - Line (12)
Call PressEnter()     ' - Line (13)

''INPUT''

'''outbound delivery

Call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_VL03N_4004_OUTBOUND_DELIVERY,False) 
Call PressEnter()
' VerifyTextBoxContent(textboxAttachedText, textboxName, textboxIndex, expectedValue, blnIsItPopup)
'VerifyTextBoxContent
' GetTableCellData(tableName, columnName, rowNumber, refColumnName, refCellValue, dataTableColumnName, blnIsItPopup)
' GetTextboxValue(textboxName, textboxIndex, dataTableColumnName, blnIsItPopup)
'call GetTextboxValue("KUWEV-KUNNR",2,"DT_ShipToParty",FALSE)
'
'call GetTableCellData("SAPMV50ATC_LIPS_OVER","Article",1,"","","DT_Article1",false)
'call GetTableCellData("SAPMV50ATC_LIPS_OVER","Article",2,"","","DT_Article2",false)
'call GetTableCellData("SAPMV50ATC_LIPS_OVER","Quantity Actually Deliveredin",1,"","","DT_QuantityActuallyDeliveredin1",false)
'call GetTableCellData("SAPMV50ATC_LIPS_OVER","Quantity Actually Deliveredin",2,"","","DT_QuantityActuallyDeliveredin2",false)
'call GetTableCellData("SAPMV50ATC_LIPS_OVER","Site",1,"","","DT_Site1",false)
'call GetTableCellData("SAPMV50ATC_LIPS_OVER","Site",2,"","","DT_Site2",false)
'call GetTableCellData("SAPMV50ATC_LIPS_OVER","SLoc",1,"","","DT_SLoc1",false)
'call GetTableCellData("SAPMV50ATC_LIPS_OVER","SLoc",2,"","","DT_SLoc2",false)
'' GetInputFromExcel(InputFilePath, sheetName, IterationIndex)
'call GetInputFromExcel(gstrInputExcelFilePathAndName,"Output",DataRowSet)
'
call VerifyTextBoxContent("Ship-To Party","KUWEV-KUNNR","2",DT_VL03N_1502_CHECK_TEXT_OF_SHIPTO_PARTY,False)
call VerifyTableCellContent(1,"Article","SAPMV50ATC_LIPS_OVER",DT_VL03N_1102_CHECK_TEXT_OF_TABLECELL_ARTICLE_0)
'Author - KGARA
'Changed the data column names according to datatsheet.
call VerifyTableCellContent(2,"Article","SAPMV50ATC_LIPS_OVER",DT_VL03N_1102_CHECK_TEXT_OF_TABLECELL_ARTICLE_1)
call VerifyTableCellContent(1,"Deliv. Qty","SAPMV50ATC_LIPS_OVER",DT_VL03N_1102_CHECK_TEXT_OF_TABLECELL_DELIV_QTY_0)
call VerifyTableCellContent(2,"Deliv. Qty","SAPMV50ATC_LIPS_OVER",DT_VL03N_1102_CHECK_TEXT_OF_TABLECELL_DELIV_QTY_1)
call VerifyTableCellContent(1,"Site","SAPMV50ATC_LIPS_OVER",DT_VL03N_1102_CHECK_TEXT_OF_TABLECELL_SITE_0)
call VerifyTableCellContent(2,"Site","SAPMV50ATC_LIPS_OVER",DT_VL03N_1102_CHECK_TEXT_OF_TABLECELL_SITE_1)
call VerifyTableCellContent(1,"SLoc","SAPMV50ATC_LIPS_OVER",DT_VL03N_1102_CHECK_TEXT_OF_TABLECELL_SLOC_0)
call VerifyTableCellContent(2,"SLoc","SAPMV50ATC_LIPS_OVER",DT_VL03N_1102_CHECK_TEXT_OF_TABLECELL_SLOC_1)
call ClickButton("Header Details   \(F8\)",false)
Call TakeScreenShot()
Call SelectTab("TAXI_TABSTRIP_HEAD","Administration",False)

call ClickButton("Document Flow   \(F7\)",false)
'''' SelectMenuItemToolBar(menuItem, toolbarIndex)
''''call SelectMenuItemToolBar("Find",2)     '' comment either sleect menu item or clickbuttontoolbar
'''''call ClickButtonToolBar("&FIND",TRUE)
'''''
'''''
''''''SetTextbox(textboxAttachedText,textboxName,textboxIndex,textboxValue,blnIsItPopup)
'''''Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_SEARCH,False)     ' - Line (19)
'''''call ClickButton("OK   \(Enter\)",false)
'''''call ClickButton("OK   \(Enter\)",false)
'''''call ClickButtonIfExist("Continue   \(Enter\)",false)
'''' ActivateItemGuiTree(treeIndex, itemPath, itemText)
''''call ActivateItemGuiTree(2,"#3;Handling unit 3534163661",)
'''call ActivateItemGuiTree(2,"Purchase Order 5522786853;Outb. del.(Intracy) 3247155119;Handling unit 3534163661","Handling unit 3534163661")

'''' ClickButtonToolBar(buttonName, toolbarIndex)
'''
'''' GetGridContent(gridTitle, gridIndex, columnName, rowNumber, refColumn, refFieldVal, dataTableColumnName)
'''call GetGridContent("Handling unit 3534163661",0,"Doc.no.",1,"NA","NA","DT_Docno_Output")

Call SelectNodeGuiTree(0,"#1;#1;RegExp:=TF .*")
'Call SelectNodeGuiTree(0,"#1;#1;RegExp:=GD .*")
Call TakeScreenShot()
Call DoubleClick()

Call GetGridContent("TF.*", 0, "Doc.no.", 1, "Item","", "DT_Docno_Output")
'Call GetGridContent("GD.*", 0, "Doc.no.", 1, "Item","", "DT_Docno_Output")
Call TakeScreenShot()

Call LogOff()
Call FinalStatus ()




