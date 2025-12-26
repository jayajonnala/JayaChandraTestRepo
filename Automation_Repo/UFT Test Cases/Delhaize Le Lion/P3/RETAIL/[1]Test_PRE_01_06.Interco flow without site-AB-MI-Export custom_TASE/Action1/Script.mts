'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_PRE_01_06.Interco flow without site-AB-MI-Export custom  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_PRE_01_06.Interco flow AB-MI-Export"
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

''--------TransactionCode-VA01----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE) 


Call SetTextbox("Order Type","VBAK-AUART","",DT_VA01_0101_ORDER_TYPE,False)
Call PressEnter()
Call TakeScreenShot
Call SetTextbox("Sold-To Party","KUAGV-KUNNR","",DT_VA01_4701_SOLDTO_PARTY,False)
Call TakeScreenShot
Call PressEnter() 
Call SelectTab("TAXI_TABSTRIP_OVERVIEW","Item overview",False)
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Article", 1, "", "", DT_VA01_4900_TABLECELL_ARTICLE_0, False)
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Article", 2, "", "", DT_VA01_4900_TABLECELL_ARTICLE_1, False)
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Article", 3, "", "", DT_VA01_4900_TABLECELL_ARTICLE_2, False)
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Order Quantity", 1, "", "", DT_VA01_4900_TABLECELL_ORDER_QUANTITY_0, False)
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Order Quantity", 2, "", "", DT_VA01_4900_TABLECELL_ORDER_QUANTITY_1, False)
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Order Quantity", 3, "", "", DT_VA01_4900_TABLECELL_ORDER_QUANTITY_2, False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetStatusBar("item2", "DT_OP_SO_NUMBER_OUTPUT")
Call GetTextStatusBar("DT_OP_VA01_4001_CHECK_TEXT_OF_STATUSBAR")

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call VerifyStatusBar(DT_VA01_4001_CHECK_TEXT_OF_STATUSBAR)


'''''''''--------TransactionCode-/nVA02----------''''

Call SetTcode(DT_VA01_4001_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)

Call SetTextbox("Order","VBAK-VBELN","",DT_VA01_0102_ORDER,False)
Call PressEnter()     
Call TakeScreenShot
'Commenting below line as we selecting only null value 
'Call SetCombo("VBAK-LIFSK"," ")
Call TakeScreenShot
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot

''''''''--------TransactionCode-/nVL10----------''''

Call SetTcode(DT_VA01_0102_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call SelectTab("TABSTRIP_ORDER_CRITERIA","Sales Orders",False)
Call TakeScreenShot
Call SetTextboxNoLabel("ST_VBELN-LOW","",DT_VA01_1020_SALES_DOCUMENT,False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot
Call SelectRowGuiGridbyRowNo("",0, 1, False)
Call ClickButtonIfExist("Create Delivery in Background   \(Shift\+F7\)",False)
Call TakeScreenShot
'Call DoubleClickGuiGridCell("","",1,"Sales Document",False)
Call DoubleClickGuiGridCell("","",1,"SD Document",False)
Call TakeScreenShot
Call GetTextboxValue("LIKP-VBELN","","DT_OP_VA01_1502_CHECK_TEXT_OF_OUTBOUND_DELIV",False)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)
Call VerifyTableCellContent(1, "Site", "SAPMV50ATC_LIPS_OVER",DT_VA01_1102_CHECK_TEXT_OF_TABLECELL_SITE_0)
Call VerifyTableCellContent(2, "Site", "SAPMV50ATC_LIPS_OVER",DT_VA01_1102_CHECK_TEXT_OF_TABLECELL_SITE_1)
Call VerifyTableCellContent(3, "Site", "SAPMV50ATC_LIPS_OVER",DT_VA01_1102_CHECK_TEXT_OF_TABLECELL_SITE_2)
Call VerifyTableCellContent(1, "SLoc", "SAPMV50ATC_LIPS_OVER",DT_VA01_1102_CHECK_TEXT_OF_TABLECELL_SLOC_0)
Call VerifyTableCellContent(2, "SLoc", "SAPMV50ATC_LIPS_OVER",DT_VA01_1102_CHECK_TEXT_OF_TABLECELL_SLOC_1)
Call VerifyTableCellContent(3, "SLoc", "SAPMV50ATC_LIPS_OVER",DT_VA01_1102_CHECK_TEXT_OF_TABLECELL_SLOC_2)
Call VerifyTableCellContent(1, "ItCa", "SAPMV50ATC_LIPS_OVER",DT_VA01_1102_CHECK_TEXT_OF_TABLECELL_ITCA_0)
Call VerifyTableCellContent(2, "ItCa", "SAPMV50ATC_LIPS_OVER",DT_VA01_1102_CHECK_TEXT_OF_TABLECELL_ITCA_1)
Call VerifyTableCellContent(3, "ItCa", "SAPMV50ATC_LIPS_OVER",DT_VA01_1102_CHECK_TEXT_OF_TABLECELL_ITCA_2)
Call SelectMenuBar("Goto;Header;Administration")
Call TakeScreenShot
Call VerifyTextBoxContent("Delivery Type", "LIKP-LFART", "", DT_VA01_2110_CHECK_TEXT_OF_DELIVERY_TYPE, False)
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
