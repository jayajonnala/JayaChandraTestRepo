
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_042-Returns to EU Vendors_P2     
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
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_P2P_01_01_042-Returns to EU Vendors_P2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_P2P_01_01_042-Returns to EU Vendors_P2.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


''''''--------TransactionCode-MIGO----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetComboByKey("GODYNPRO-ACTION", DT_SETKEY_1)
Call SetComboByKey("GODYNPRO-REFDOC", DT_SETKEY_2)
Call SetTextbox("GR goods receipt","GODEFAULT_TV-BWART","",DT_MIGO_0010_GODEFAULT_TVBWART,False)
Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER", 0, DT_MIGO_2000_GODYNPROPO_NUMBER, False)
Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_MIGO_0110_DELIVERY_NOTE,False)
Call TakeScreenShot
Call PressEnter() 
Call PressEnter() 
Call ClickButtonIfExist("Open detail data", False)
Call SelectTab("TS_GOITEM","Where",False)
Call TakeScreenShot
Call VerifyTextBoxContent("Movement type","GOITEM-BWART","",DT_MIGO_0325_CHECK_TEXT_OF_MOVEMENT_TYPE,False)
Call SelectTab("TS_GOITEM","Quantity",False)
Call TakeScreenShot
Call SetTextbox("Qty in Unit of Entry","GOITEM-ERFMG","",DT_MIGO_0300_QUANTITY,False)
Call PressEnter() 
Call TakeScreenShot
Call SelectCheckbox("GODYNPRO-DETAIL_TAKE", 1, "ON", False)
Call SetTableData("SAPLMIGOTV_GOITEM","OK","2","","","ON",False) 
Call PressEnter() 

Call ClickButton("Check Entries   \(F7\)",false)
Call TakeScreenShot
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButton("Post Document   \(Shift\+F11\)",false)
Call TakeScreenShot
Call GetStatusBar("item1","DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Article document "&DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" posted")
Call SetComboByKey("GODYNPRO-ACTION", DT_MIGO_0010_GODYNPROACTION)
Call SetTextboxNoLabel("GODYNPRO-MAT_DOC", "", DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT, False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

'''''--------TransactionCode-MB51----------''''

Call SetTcode(DT_MIGO_0001_OKCD)     
Call PressEnter()     

Call SetTextbox("Article Document","MBLNR-LOW","",DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT,False)
Call SetTextbox("Company Code","BUKRS-LOW","",DT_MIGO_1000_COMPANY_CODE,False)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",false)
Call TakeScreenShot
Call ClickButtonIfExist("Detail List   \(Ctrl\+Shift\+F12\)",False)
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "BWART", 0, DT_MIGO_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWART)
Call VerifyGridCellContent("", 1, "MBLNR", 0, DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT)
Call VerifyGridCellContent("", 2, "MBLNR", 0, DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT)
Call VerifyGridCellContent("", 1, "EBELN", 0, DT_MIGO_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_EBELN)
Call VerifyGridCellContent("", 2, "EBELN", 0, DT_MIGO_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_EBELN)
Call SelectRowGuiGrid("", 0, "Article Document", DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT, False)
Call DoubleClickGuiGridCell("", 0, 1, "Article Document", False)
Call TakeScreenShot
Call SelectTab("TS_GOITEM","Output",False)
Call TakeScreenShot
Call ClickButton("Display outputs",false)
Call TakeScreenShot
Call VerifyTableCellContent(1, "Status", "SAPDV70ATC_NAST3", DT_MIGO_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_0)
Call VerifyTableCellContent(1, "Status", "SAPDV70ATC_NAST3", DT_MIGO_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_0)
Call VerifyTableCellContent(1, "Output Type", "SAPDV70ATC_NAST3", DT_MIGO_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_0)
Call VerifyTableCellContent(1, "Output Type", "SAPDV70ATC_NAST3", DT_MIGO_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_0)
Call VerifyTableCellContent(1, "Medium", "SAPDV70ATC_NAST3", lcase(DT_MIGO_0100_CHECK_VALUE_OF_TABLECELL_MEDIUM_0))
Call VerifyTableCellContent(1, "Medium", "SAPDV70ATC_NAST3", lcase(DT_MIGO_0100_CHECK_VALUE_OF_TABLECELL_MEDIUM_0))

''''''--------TransactionCode-MB90----------''''

Call SetTcode(DT_MIGO_0100_OKCD)     
Call PressEnter()     

Call SetTextbox("Article Document","RG_MBLNR-LOW","",DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT,False)
Call SetTextbox("Article Doc. Year","PM_MJAHR","",year(Date),False)
Call SetTextbox("Processing Mode","PM_VERMO","",DT_MIGO_1000_PROCESSING_MODE,False)
Call SetTextbox("Sort order","PM_NSORT","",DT_MIGO_1000_SORT_ORDER,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call SelectCheckboxNoLabel(0, DT_MIGO_0120_NO_NAME, False)
Call ClickButton("Print preview   \(Shift\+F4\)",False)
Call TakeScreenShot

''''''--------TransactionCode-MB03----------''''

Call SetTcode(DT_MIGO_0750_OKCD)     
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Article Doc.","RM07M-MBLNR","",DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT,False)
Call SetTextbox("Art\. Doc\. Year","RM07M-MJAHR","",year(Date),False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Accounting Documents   \(F7\)",False)
Call TakeScreenShot
Call GetTextboxValue("BKPF-BELNR",0,"DT_MIGO_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OUTPUT", False)


''''''--------TransactionCode-FB03----------''''
Call SetTcode(DT_MIGO_0750_OKCD_OCC1)     
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Document Number","RF05L-BELNR","",DT_MIGO_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OUTPUT,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_MIGO_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",year(Date),False)
Call PressEnter()
Call TakeScreenShot
Call VerifyTextBoxContent("Document Number", "BKPF-BELNR", 0, DT_MIGO_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OUTPUT, False)
Call VerifyGridCellContent("", 1, "BUKRS", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUKRS)
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)

''''''--------TransactionCode-/nZMDPU_MB5S----------''''

Call SetTcode(DT_MIGO_0750_OKCD_OCC2)     
Call PressEnter()

Call SetTextboxNoLabel("LIFNR-LOW","",DT_MIGO_1000_VENDOR,False)
Call SetTextbox("Purch. Organization","EKORG-LOW","",DT_MIGO_1000_PURCH_ORGANIZATION,False)
Call SetTextbox("Movement type","S_BWART-LOW","",DT_MIGO_1000_MOVEMENT_TYPE,False)
Call SetTextbox("Partner Function","S_PARVW-LOW","",DT_MIGO_1000_PARTNER_FUNCTION,False)
Call SetTextbox("Site","S_WERKS-LOW","",DT_Site,False)
Call SetTextbox("Entry Date","S_EDATE-LOW","",ConvertDate(DT_Entry_Date),False)
Call SetTextboxNoLabel("S_BUKRS-LOW",0,DT_COMPANY_CODE,False)
Call TakeScreenShot
Call PressEnter()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call SelectCheckboxNoLabel(0,DT_MIGO_0120_NO_NAME_OCC1,False)
Call SelectCheckboxNoLabel(1,DT_MIGO_0120_NO_NAME_OCC2,False)
Call ClickButton("Send notification   \(Shift\+F2\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)



Call LogOff()

Call FinalStatus ()






 

