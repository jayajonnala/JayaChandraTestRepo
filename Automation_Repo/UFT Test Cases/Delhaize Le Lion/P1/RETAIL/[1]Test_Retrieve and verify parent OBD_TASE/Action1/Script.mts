
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


gstrTestCaseName = "Test_Retrieve and verify parent OBD_TASE"'''' "Test_Retr and ver parent OBD"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'-------------------------------------ME23N------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter()
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call SetTextbox("Pur. Order","MEPO_SELECT-EBELN","",DT_ME23N_0003_PUR_ORDER,True)  
Call TakeScreenShot
Call SelectRadioButton("MEPO_SELECT-BSTYP_F","Pur. Order",True)     
Call ClickButton("Other Document   \(Enter\)",True)  
Call TakeScreenShot
Call ClickButtonIfExist("Expand Item Details Ctrl\+F4",False)
Call SelectTab("ITEM_DETAIL","Purchase Order History",False) 
Call TakeScreenShot
'Call GetGridContent("",0,"Article Document",1,"<NA>","<NA>","DT_ME23N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR")
'Call TakeScreenShot
''Call ClickContextButtonToolBar("FIND",0)

''Call GetGridContentByRefColumn("","","Article Document",DT_ME23N_0841_SEARCH_TERM,"Short Text",DT_ME23N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
''Call SetTextbox("Search Term","GS_SEARCH-VALUE","",DT_ME23N_0841_SEARCH_TERM,True)
''Call ClickButtonIfExist("OK   \(Enter\)",True)
'SAPGuiSession("Session_2").SAPGuiWindow("STO Retail 5542350983").Maximize
'SAPGuiSession("Session_2").SAPGuiWindow("STO Retail 5542350983").SAPGuiTabStrip("ITEM_DETAIL").Select "Material Data"
'SAPGuiSession("Session_2").SAPGuiWindow("STO Retail 5542350983").SAPGuiTabStrip("ITEM_DETAIL").Select "Purchase Order History"
'SAPGuiSession("Session_2").SAPGuiWindow("STO Retail 5542350983").SAPGuiToolbar("GridToolbar").PressButton "&FIND"
'SAPGuiSession("Session_2").SAPGuiWindow("Find").SAPGuiEdit("Search Term:").Set "Lfs"
'SAPGuiSession("Session_2").SAPGuiWindow("Find").SAPGuiButton("OK   (Enter)").Click
'SAPGuiSession("Session_2").SAPGuiWindow("Find").SAPGuiEdit("Information Message").SetFocus
'SAPGuiSession("Session_2").SAPGuiWindow("Find").SAPGuiButton("OK   (Enter)").Click
'SAPGuiSession("Session_2").SAPGuiWindow("Find").SAPGuiButton("Cancel   (F12)").Click
''Call ClickButtonIfExist("Cancel   \(F12\)",True)


call ClickCellGuiGrid("",1,"Article Document","","Short Text","Lfs",False)

Call GetTextBoxValue("LIKP-VBELN",0,"DT_ME23N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_OUTPUT",False)
Call SelectTab("TAXI_TABSTRIP_.*","Item Overview",False)
call SelectRowGuiTable("SAPMV50ATC_LIPS_.*","Itm",10,False)
Call TakeScreenShot



Call ClickButton("BT_IDET_T",False)  ' click item detail 
Call SelectTab("TAXI_TABSTRIP_.*","Administration",False) 
Call VerifyTextBoxContent("Reason for mvt\.", "LIPS-GRUND", 8, DT_ME23N_0100_GRIDCELL_0_ARTICLE_DOCUMENT_OCC1, False)
Call TakeScreenShot

Call ClickButton("btn\[3\]",False)

Call SelectTab("TAXI_TABSTRIP_.*","Item Overview",False)
call SelectRowGuiTable("SAPMV50ATC_LIPS_.*","Itm",20,False)
Call TakeScreenShot
Call ClickButton("BT_IDET_T",False)
Call SelectTab("TAXI_TABSTRIP_.*","Administration",False) 
Call VerifyTextBoxContent("Reason for mvt\.", "LIPS-GRUND", 8, DT_ME23N_0100_GRIDCELL_0_ARTICLE_DOCUMENT_OCC1, False)
Call TakeScreenShot
Call ClickButton("btn\[3\]",False)

''''2nd row'''''
'''1st row selection ''''
Call SelectTab("TAXI_TABSTRIP_.*","Item Overview",False)
call SelectRowGuiTable("SAPMV50ATC_LIPS_.*","Itm",10,False)
Call TakeScreenShot
Call ClickButton("BT_IDET_T",False)

Call SelectTab("TAXI_TABSTRIP_.*","Picking",False)

Call GetTextBoxValue("LIPS-LFIMG", 5, "DT_PICK_QTY_1_OUTPUT", False)
Call GetTextBoxValue("LIPSD-PIKMG", 5, "DT_PICK_QTY_2_OUTPUT", False)
Call GetTextBoxValue("LIPS-LGMNG", 5, "DT_PICK_QTY_3_OUTPUT", False)
Call GetTextBoxValue("LIPS-UMVKN",5,"DT_PICK_QTY_4_OUTPUT", False)
Call TakeScreenShot
Call ClickButton("btn\[3\]",False)

Call SelectTab("TAXI_TABSTRIP_.*","Item Overview",False)
call SelectRowGuiTable("SAPMV50ATC_LIPS_.*","Itm",20,False)
Call TakeScreenShot
Call ClickButton("BT_IDET_T",False)
Call SelectTab("TAXI_TABSTRIP_.*","Picking",False)
Call TakeScreenShot
Call GetTextBoxValue("LIPS-LFIMG", 5, "DT_DELIV_QTY_1_ART_2_OUTPUT", False)
Call GetTextBoxValue("LIPSD-PIKMG", 5, "DT_DELIV_QTY_2_ART_2_OUTPUT", False)
Call GetTextBoxValue("LIPS-LGMNG", 5, "DT_DELIV_QTY_3_ART_2_OUTPUT", False)
Call GetTextBoxValue("LIPS-UMVKN",5,"DT_DELIV_QTY_4_ART_2_OUTPUT", False)
Call TakeScreenShot
Call ClickButton("btn\[3\]",False)
Call VerifyTextBoxContent("Ship-To Party", "KUWEV-KUNNR", 1, DT_ME23N_1502_CHECK_TEXT_OF_SHIPTO_PARTY, False)
Call SelectTab("TAXI_TABSTRIP_.*","Picking",False)
'Call VerifyTextBoxContent("Pick Date/Time", "LIKP-KODAT", 4, ConvertDate(DT_ME23N_1104_CHECK_TEXT_OF_PICK_DATETIME), False)
Call TakeScreenShot
Call SelectTab("TAXI_TABSTRIP_.*","Item Overview",False)
call VerifyTableCellContent(1,"Article","SAPMV50ATC_LIPS_OVER",DT_ME23N_1102_CHECK_TEXT_OF_TABLECELL_ARTICLE_0)
call VerifyTableCellContent(2,"Article","SAPMV50ATC_LIPS_OVER",DT_ME23N_1102_CHECK_TEXT_OF_TABLECELL_ARTICLE_1)
call VerifyTableCellContent(1,"Deliv. Qty","SAPMV50ATC_LIPS_OVER",DT_ME23N_1102_CHECK_TEXT_OF_TABLECELL_DELIV_QTY_0)
call VerifyTableCellContent(2,"Deliv. Qty","SAPMV50ATC_LIPS_OVER",DT_ME23N_1102_CHECK_TEXT_OF_TABLECELL_DELIV_QTY_1)
call VerifyTableCellContent(1,"Site","SAPMV50ATC_LIPS_OVER",DT_ME23N_1102_CHECK_TEXT_OF_TABLECELL_SITE_0)
call VerifyTableCellContent(2,"Site","SAPMV50ATC_LIPS_OVER",DT_ME23N_1102_CHECK_TEXT_OF_TABLECELL_SITE_1)
call VerifyTableCellContent(1,"SLoc","SAPMV50ATC_LIPS_OVER",DT_ME23N_1102_CHECK_TEXT_OF_TABLECELL_SLOC_0)
call VerifyTableCellContent(2,"SLoc","SAPMV50ATC_LIPS_OVER",DT_ME23N_1102_CHECK_TEXT_OF_TABLECELL_SLOC_1)
Call TakeScreenShot

Call LogOff()
Call FinalStatus ()


