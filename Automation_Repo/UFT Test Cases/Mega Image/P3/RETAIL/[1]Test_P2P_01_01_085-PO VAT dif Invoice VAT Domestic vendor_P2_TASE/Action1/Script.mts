'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_P2P_01_01_085-PO VAT dif Invoice VAT Domestic vendor_P2_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 2nd June
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_P2P_01_01_085_P2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_P2P_01_01_085-PO VAT  dif Invoice VAT  Domestic vendor_P2_TASE.xls"
'''''----------------------Login----------------------------

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

'
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''
''''----------------------Tcode MIGO----------------------------
''Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()
'
'Enter the PO Number and Press Enter
'Call SetCombo("GODYNPRO-ACTION","Goods Receipt")
'Call SetCombo("GODYNPRO-REFDOC","Purchase Order")
Call SetComboByKey("GODYNPRO-ACTION",DT_MIGO_0010_GODYNPROACTION)
'Capture the screenshot
Call TakeScreenShot()

Call SetComboByKey("GODYNPRO-REFDOC",DT_MIGO_0010_GODYNPROREFDOC)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER",0,DT_MIGO_2000_GODYNPROPO_NUMBER,False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_MIGO_0110_DELIVERY_NOTE,False)
Call SetTextbox("Document Date","GOHEAD-BLDAT","",Replace((DT_MIGO_0110_DOCUMENT_DATE),"/","."),False) 
Call SetTextbox("Posting Date","GOHEAD-BUDAT","",Replace((DT_MIGO_0110_POSTING_DATE),"/","."),False) 
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()
Call SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,"ON",False)
Call SetTableDataNoRef("SAPLMIGOTV_GOITEM","OK",1,"ON",False)
Call SetTableDataNoRef("SAPLMIGOTV_GOITEM","OK",2,"ON",False)
Call SetTableDataNoRef("SAPLMIGOTV_GOITEM","OK",3,"ON",False)
Call SetTableDataNoRef("SAPLMIGOTV_GOITEM","OK",4,"ON",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Check Entries   \(F7\)",False)
Call VerifyStatusBar(DT_MIGO_0001_GET_TEXT_OF_STATUSBAR)
'Capture the screenshot
Call TakeScreenShot()

'Post the Article No
Call ClickButton("Post   \(Ctrl\+S\)",False)
Wait(2)
Call ClickButtonIfExist("Save",True) 
Wait(2)
'Capture the screenshot
Call TakeScreenShot()
'save it to data sheet
Call GetStatusBar("item1","DT_DOCUMENT_NO_OUTPUT")
'reload data sheet
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'verify statusbar with datasheet feed
Call VerifyStatusBar(DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR_OCC3)
Call VerifyStatusBarMessageType("S")

'_______________________________________________________________________________________________________

Call SetComboByKey("GODYNPRO-ACTION",DT_MIGO_0010_GODYNPROACTION1)
'Capture the screenshot
Call TakeScreenShot()

Call SetComboByKey("GODYNPRO-REFDOC",DT_MIGO_0010_GODYNPROREFDOC1)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextboxNoLabel("GODYNPRO-MAT_DOC",0,DT_MIGO_2010_GODYNPROMAT_DOC,False)
Call SetTextboxNoLabel("GODYNPRO-DOC_YEAR",0,DT_MIGO_2010_FISCAL_YEAR,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()
Call PressEnter()

'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Open detail data",False) 
Wait(2)
Call SelectTab("TS_GOITEM","Purchase Order Data",False)
'Capture the screenshot
Call TakeScreenShot()
Call VerifyTextBoxContent("Purchase order","GOITEM-EBELN","",DT_MIGO_0320_CHECK_TEXT_OF_PURCHASE_ORDER,False)
Call VerifyTextBoxContent("Reference Document","GOITEM-LFBNR","",DT_MIGO_0320_CHECK_TEXT_OF_REFERENCE_DOCUMENT,False)

Call SelectTab("TS_GOITEM","Where",False)
'Capture the screenshot
Call TakeScreenShot()
Call VerifyTextBoxContent("Site","GOITEM-WERKS","",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_SITE_0,False)
Call ClickButton("Next Item",False) 
wait(2)
Call VerifyTextBoxContent("Site","GOITEM-WERKS","",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_SITE_1,False)
Call ClickButton("Next Item",False) 
wait(2)
Call VerifyTextBoxContent("Site","GOITEM-WERKS","",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_SITE_2,False)
Call ClickButton("Next Item",False) 
wait(2)
Call VerifyTextBoxContent("Site","GOITEM-WERKS","",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_SITE_4,False)

Call SelectTab("TS_GOITEM","Partner",False)
'Capture the screenshot
Call TakeScreenShot()
Call VerifyTextBoxContent("Vendor","GOITEM-LIFNR","",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_VENDOR_4,False)
Call ClickButton("Previous Item",False) 
wait(2)
Call VerifyTextBoxContent("Vendor","GOITEM-LIFNR","",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_VENDOR_2,False)
Call ClickButton("Previous Item",False) 
wait(2)
Call VerifyTextBoxContent("Vendor","GOITEM-LIFNR","",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_VENDOR_1,False)
Call ClickButton("Previous Item",False) 
wait(2)
Call VerifyTextBoxContent("Vendor","GOITEM-LIFNR","",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_VENDOR_0,False)

Call VerifyTableCellContent(1,"Profit Center","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_PROFIT_CENTER_0)
Call VerifyTableCellContent(2,"Profit Center","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_PROFIT_CENTER_1)
Call VerifyTableCellContent(3,"Profit Center","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_PROFIT_CENTER_2)
Call VerifyTableCellContent(4,"Profit Center","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_PROFIT_CENTER_4)

Call VerifyTableCellContent(1,"Article","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_ARTICLE_0)
Call VerifyTableCellContent(2,"Article","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_ARTICLE_1)
Call VerifyTableCellContent(3,"Article","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_ARTICLE_2)
Call VerifyTableCellContent(4,"Article","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_ARTICLE_4)

Call VerifyTableCellContent(1,"Qty in UnE","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_ORDER_QUANTITY_0)
Call VerifyTableCellContent(2,"Qty in UnE","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_ORDER_QUANTITY_1)
Call VerifyTableCellContent(3,"Qty in UnE","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_ORDER_QUANTITY_2)
Call VerifyTableCellContent(4,"Qty in UnE","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_ORDER_QUANTITY_4)

Call VerifyTableCellContent(1,"Received","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_RECEIVED_0)
Call VerifyTableCellContent(2,"Received","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_RECEIVED_1)
Call VerifyTableCellContent(3,"Received","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_RECEIVED_2)
Call VerifyTableCellContent(4,"Received","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_RECEIVED_4)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()


