'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_02 Non planogram listing
'.................Test Scenario: AT_02 Non planogram listing
'.................TCode: MM42;WSM4L;ZMDAM_UPD_BLOCK
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

gstrTestCaseName = "Test_02 Non planogram listing"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.

'//----------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------//

''Call CloseSessionsSAP()
'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)'.......................Mandatory Initial Call only in First Component in a Test Scenario
'
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'TCode
Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter()
Call SetTextbox("Article","RMMW1-MATNR","",DT_MM42_0100_ARTICLE,False)
Call SetTextbox("Purchasing Org.","RMMW1-EKORG","",DT_MM42_0100_PURCHASING_ORG,False)
'Call SetTextbox("Vendor","RMMW1-LIFNR","",DT_MM42_0100_VENDOR,False)
Call SetTextboxNoLabel("RMMW1-LIFNR","",DT_MM42_0100_VENDOR,False)
Call TakeScreenShot()

Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100", "Screen description", "Basic Data", False)

Call PressEnter()

'Call ClickButton("Go to additional data   (Ctrl+F6)",False)

Call ClickButton("Go to additional data   \(Ctrl\+F6\)",False)
Call TakeScreenShot()

'SAPGuiSession("Session").SAPGuiWindow("Change Article 7176413").SAPGuiTabStrip("TABSPR1").Select "Layout Modules"
Call SelectTab("TABSPR1","Layout Modules", false)
Call TakeScreenShot()

''Call VerifyTableCellContent(1,"Layout Module","SAPLWLAY_DIA_ARTTC_LAY",DT_MM42_8024_TABLECELL_LAYOUT_MODULE_0)
Call SetTableDataNoRef("SAPLWLAY_DIA_ARTTC_LAY","Layout Module",1,DT_MM42_8024_TABLECELL_LAYOUT_MODULE_0, False)
Call SetTableDataNoRef("SAPLWLAY_DIA_ARTTC_LAY","Facing",1,DT_MM42_8024_TABLECELL_FACING_0, False)

'SetHorizontalScrollBar(10,False)
Call PressEnter()
Call TakeScreenShot()
Call ClickButton("Go to main data   \(Ctrl\+Shift\+F3\)",False)
Call ClickButton("Save   \(Ctrl\+S\)",False)
'enable status verify
'Call VerifyStatusBar("Article "&DT_MM42_0100_ARTICLE&" changed")
Call VerifyStatusBar(DT_MM42_0100_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot()

'If SAPGuiSession("Session").SAPGuiWindow("Change Article 7176413_3").SAPGuiTable("SAPLWLAY_DIA_ARTTC_LAY").GetCellData ("1","Layout Module") ="RS_NON_PLG" Then
'SAPGuiSession("Session").SAPGuiWindow("Change Article 7176413_3").SAPGuiTable("SAPLWLAY_DIA_ARTTC_LAY").SetCellData "1","Facing",DT_MM42_8024_TABLECELL_FACING_0
'End If
'SAPGuiSession("Session").SAPGuiWindow("Change Article 7176413_3").SAPGuiButton("Main data").Click
'SAPGuiSession("Session").SAPGuiWindow("Change Article 7176413").SAPGuiButton("Save   (Ctrl+S)").Click


Call SetTcode(DT_EXPECTEDTRANSACTIONCODE_OCC1)
Call PressEnter()

'SAPGuiSession("Session").SAPGuiWindow("Automatic Relisting by").SAPGuiRadioButton("Documents for a Date").Set
Call SelectRadioButtonIfExist("S_DATUM","",False)
Call SetTextbox("Assortment","FILIA-LOW","",DT_MM42_8024_TABLECELL_LAYOUT_MODULE_0,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)", False)
Wait(30)

Call SetTcode(DT_EXPECTEDTRANSACTIONCODE_OCC2)
Call TakeScreenShot()
Call PressEnter()
Call SetTextbox("Article","S_MATNR-LOW","",DT_MM42_0100_ARTICLE,False)
Call SetTextbox("Site","S_WERKS-LOW","",DT_MM42_1000_SITE,False)
Call ClickButton("Execute   \(F8\)", False)
Call SetTextbox("P-S artl status","MARC-MMSTA","",DT_MM42_0100_PS_ARTL_STATUS,False)

Call SetGridCheckbox(1,"Checkbox","ON",False)
Call SetGridCheckbox(3,"Checkbox","ON",False)
Call SetGridCheckbox(5,"Checkbox","ON",False)
Call SetGridCheckbox(7,"Checkbox","ON",False)
Call SetGridCheckbox(9,"Checkbox","ON",False)
Call SetGridCheckbox(11,"Checkbox","ON",False)
Call SetGridCheckbox(13,"Checkbox","ON",False)
Call SetGridCheckbox(15,"Checkbox","ON",False)
Call SetGridCheckbox(17,"Checkbox","ON",False)
Call SetGridCheckbox(19,"Checkbox","ON",False)
Call SetGridCheckbox(21,"Checkbox","ON",False)
Call SetGridCheckbox(23,"Checkbox","ON",False)
Call SetGridCheckbox(25,"Checkbox","ON",False)

'SAPGuiSession("Session").SAPGuiWindow("Update blocking status").SAPGuiButton("Save   (Ctrl+S)").Click
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButton("BUTTON_2",False)
'SAPGuiSession("Session").SAPGuiWindow("Update blocking status_2").SAPGuiButton("Background").Click


Call LogOff()
Call FinalStatus()

