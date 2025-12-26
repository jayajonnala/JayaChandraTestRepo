

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Article Creation via DAP_P2
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

gstrTestCaseName = "Test_Article Creation via DAP_P2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_002_Create_Purchasing_Conditions_for_an_Existing_Article_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'''''''''''''''''''''''''''''''''''------------------------------------------'''''''''''''''''''''''''''''''''''''
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


'----------------------Tcode CV04N ----------------------------
'Enter the transaction code
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenSHot()


Call SelectTab("MAINSTRIP","Classification",False)
Call TakeScreenShot()

Call SetTextbox("Class","MCDOK-KLASSE","",DT_CV04N_0403_CLASS,False)
Call SetTextbox("Class Type","MCDOK-KLASSENTYP","",DT_CV04N_0403_CLASS_TYPE,False)
Call TakeScreenShot()
Call PressEnter()  
Call TakeScreenShot()
Wait 2
Call SetTextbox("Article Description","RCTMS-MWERT","",DT_CV04N_4000_RCTMSMWERT,False)
Call SetTextbox("Source System","RCTMS-MWERT","","DAP",False)
Call ClickButton("Execute   \(F8\)",False)
'Added wait time due to page laoding taking time
Wait 20
Call GetGridContentByTitle("", 0, "DOKNR", 1, "DT_CV04N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DOKNR_OUTPUT")
Call GetGridContentByTitle("", 0, "Global Trade Item Number GTIN", 1, "DT_CV04N_0500_GTIN_BASE_UOM_OUTPUT")

Call SelectRowGuiGridbyRowNo("", 0, 1, False)
Call TakeScreenShot()
Call ClickButton("Display   \(Ctrl\+Shift\+F9\)",False)
Call TakeScreenShot()
Call ClickButton("Display <-> Change   \(Shift\+F8\)",False)
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call VerifyStatusBar("Document info record ZAC "&DT_CV04N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DOKNR_OUTPUT&" 000 00 changed")

Call ClickButton("Back   \(F3\)",False)
Wait 2
Call ClickButton("Execute   \(F8\)",False)

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call VerifyGridCellContent("", 1, "DOKNR", 0, DT_CV04N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DOKNR_OCC1)
Call VerifyGridCellContent("", 1, "STATUSTEXT", 0, LCase(DT_CV04N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_STATUSTEXT))

''----------------------Tcode W_SYNC ----------------------------
'Enter the transaction code
Call SetTcode(DT_CV04N_0500_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_CV04N_0500_OKCD)
Call TakeScreenSHot()


Call SelectCheckbox("P_MYITEM",0,"OFF",False)

Call SetTextbox("Profile","P_PROFIL","",DT_CV04N_0500_PROFILE,False)
Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("GTIN Base UoM","PI_06-LOW","",DT_CV04N_0500_GTIN_BASE_UOM,False)

'Click on Select Button
Call ClickButton("Select   \(F8\)",False)

'Capture the screenshot
Call TakeScreenShot()

Call ActivateNodeGuiTree(0,"#1;#1;#1")
Wait(2)
Call TakeScreenShot()

'Verify the grid content
Call VerifyGridCellContent("",1,"GTIN Base UoM",0,DT_CV04N_5000_CHECK_GETCELLVALUE_OF_GRIDCELL_0_GTIN)
Call VerifyGridCellContent("",1,"Processing Status",0,DT_CV04N_5000_CHECK_GETCELLVALUE_OF_GRIDCELL_0_GDSPROC_STATUS)

'Get The Article Description
Call GetGridContent("",0,"Object Description",1,"GTIN Base UoM",DT_CV04N_5000_CHECK_GETCELLVALUE_OF_GRIDCELL_0_GTIN,"DT_CV04N_5000_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DESCRIPTION_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'Click on Link to Price Catalog Maintenance
Call SelectRowGuiGrid("",0,"GTIN Base UoM",DT_CV04N_0500_GTIN_BASE_UOM,False)
Call ClickButton("Link to Price Catalog Maintenance   \(F7\)",False)

'Select Row
Call SelectRowGuiGrid("",0,"Processing Status","01",False)
Wait(1)

'Click on Create/Change
Call ClickButton("Create/Change Article - Immediately   \(Shift\+F9\)",False)
Wait(2)

Call VerifyGridCellContent("",5,"Processing Status",0,DT_CV04N_0400_CHECK_GETCELLVALUE_OF_GRIDCELL_4_PROC_STATUS)
Call VerifyGridCellContent("",4,"GTIN Base UoM",0,DT_CV04N_0500_GTIN_BASE_UOM)
Call VerifyGridCellContent("",4,"Artl Short Text ",0,DT_CV04N_4000_RCTMSMWERT)
Call VerifyGridCellContent("",4,"Material Type",0,DT_CV04N_0400_CHECK_GETCELLVALUE_OF_GRIDCELL_4_MATL_TYPE)



Call SelectRowGuiGrid("",0,"Processing Status","03",False)
Wait(1)

'Refresh the screen
Call ClickButtonToolBar("&REFRESH",0)
Wait(5)
'Refresh the screen
Call ClickButtonToolBar("&REFRESH",0)
Wait(5)
Call ClickButtonToolBar("&REFRESH",0)
Wait(5)
'Refresh the screen
Wait(5)
'Refresh the screen
Call ClickButtonToolBar("&REFRESH",0)
Wait(5)
Call ClickButtonToolBar("&REFRESH",0)
Wait(5)

Call VerifyGridCellContent("",6,"Processing Status",0,DT_CV04N_0400_CHECK_GETCELLVALUE_OF_GRIDCELL_4_PROC_STATUS_OCC1)
Call VerifyGridCellContent("",4,"Artl Short Text ",0,DT_CV04N_4000_RCTMSMWERT)
Call VerifyGridCellContent("",4,"Material Type",0,DT_CV04N_0400_CHECK_GETCELLVALUE_OF_GRIDCELL_4_MATL_TYPE_OCC1)

Call GetGridContentByRefColumn("",0,"Processing Status","04","Article","DT_NEW_ARTICLE_OUTPUT")

Call WriteRunTimeDataToExcelGlobalSheet ("DT_NEW_ARTICLE_OUTPUT",DT_NEW_ARTICLE)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call FindRowNumber("","Processing Status", "04", "DT_ROWNUMBER")

Call ClickCellGuiGrid("", 0, "Article",DT_ROWNUMBER, "Processing Status", "04", False)
Call TakeScreenshot()

'''''''---------MM42-----------------------'''''''

Call SetTextbox("Purchasing Org\.","RMMW1-EKORG","",DT_PURCHASING_ORG,False)
Call SetTextbox("Sales Org\.","RMMW1-VKORG","",DT_SALES_ORG,False)
Call SetTextbox("Distr\. Channel","RMMW1-VTWEG","",DT_DISTR_CHANNEL,False)
Call TakeScreenshot()

Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100", "Screen description","Basic Data", False)
Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100", "Screen description","Purchasing", False)
Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100", "Screen description","Sales", False)
Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100", "Screen description","Logistics: Distribution Center", False)
Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100", "Screen description","Logistics: Store", False)
Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100", "Screen description","POS", False)
Call TakeScreenshot()
Call PressEnter()

Call SetVerticalScrollBar(100,False)
Call TakeScreenshot()
Call VerifyTextBoxContent("Brand", "MARA-ZZBRAND", 0, DT_BRAND, False)

Call ClickButton("Display item details",False)
Call TakeScreenshot()

Call ClickCellTableByRowNo("SAPLZMDAM_EXT_SCRTC_LTEXT", "#1", 1, False)
Call TakeScreenshot()
Call ClickButton("Create text",False)
Call TakeScreenshot()
Call SetCombo("Create text in",DT_TEXT)
Call TakeScreenshot()
Call ClickButton("Enter   \(Enter\)",True)
Call SetTextArea("Automation Testing")

Call ClickButton("Create text",False)
Call SetComboByKey("Create text in", "G")
Call TakeScreenshot()
Call ClickButton("Enter   \(Enter\)",True)
Call SetTextArea("Automation Testing")
Call TakeScreenshot()

Call ClickBUtton("Back   \(F3\)",False)

Call SelectTab("TABSPR1", " Purchasing", False)
Call TakeScreenshot()
Call ClickButton("InfoRec\.Texts",False)
Call TakeScreenshot()

Call SetTextbox("Text","RM06I-LTEX1","1",DT_ARTICLE_FORM_DESCRIPTION,False)
Call SetTextbox("Text","RM06I-LTEX2","1",DT_PURCHASE_ORDER_TEXT_OCC1,False)
Call TakeScreenshot()
Call SelectMenuBar("Edit;Texts;Other Languages...")
Call TakeScreenshot()
Call SetTextbox("Next language","RM06I-SPRA2","","EN",True)
Call TakeScreenshot()
Call ClickButton("Continue   \(Enter\)",False)
Call TakeScreenshot()
Call SetTextbox("Text","RM06I-LTEX1","1",DT_ARTICLE_FORM_DESCRIPTION,False)
Call SetTextbox("Text","RM06I-LTEX2","1",DT_PURCHASE_ORDER_TEXT_OCC1,False)
Call TakeScreenshot()

Call ClickButton("Back   \(F3\)",False)
Call SelectTab("TABSPR1", " Sales", False)
Call SetVerticalScrollBar(100,False)
Call TakeScreenshot()
Call VerifyTextBoxContent("ArticleCmpGroup","MVKE-MVGR1", 0, DT_ArticleCmpGroup, False)
 Call TakeScreenshot()
Call SelectTab("TABSPR1", " Logistics: store", False)
Call SetVerticalScrollBar(100,False)
Call TakeScreenshot()
Call ClickBUtton("Other Logistics Data",False)
Call VerifyTextBoxContent("Max\. Storage Period","MARC-MAXLZ", 0, DT_MX_STORAGE_PERIOD, False)
Call ClickBUtton("Go to main data   \(Ctrl\+Shift\+F3\)",False)
Call TakeScreenshot()

Call TakeScreenshot()
Call ClickButton("Go to additional data   \(Ctrl\+F6\)",False)

Call SelectTab("TABSPR1", "Additional GTINs", False)
Call TakeScreenshot()
CAll FindRowNumber("SAPLMGD2TC_EAN", "Unit text", "each", "DT_ROW_NUMBER")
CAll VerifyTableCellContent(DT_ROW_NUMBER, "GTIN", "SAPLMGD2TC_EAN", DT_CV04N_0500_GTIN_BASE_UOM)
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenshot()
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call VerifyStatusBarMessageType("S")
Call VerifyStatusBar("Article "&DT_NEW_ARTICLE&" Changed")

Call FindRowNumber("","Processing Status", "04", "DT_ROWNUMBER")
Call ClickCellGuiGrid("", 0, "Article",DT_ROWNUMBER, "Processing Status", "04", False)
Call TakeScreenshot()

Call ClickButton("Deselect All   \(Shift\+F7\)",False)
Call TakeScreenshot()

Call SetTextbox("Purchasing Org\.","RMMW1-EKORG","",DT_PURCHASING_ORG,False)
Call SetTextbox("Sales Org\.","RMMW1-VKORG","",DT_SALES_ORG,False)
Call SetTextbox("Distr\. Channel","RMMW1-VTWEG","",DT_DISTR_CHANNEL_OCC1,False)
Call TakeScreenshot()

Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100", "Screen description","Sales", False)
Call TakeScreenshot()
Call PressEnter()
Call TakeScreenshot()
CAll SetVerticalScrollBar(100, False)
Call VerifyTextBoxContent("Item Set","MVKE-MVGR4", 0, DT_ITEMSET, False)


'''''''----------------------Tcode  VKP5 ----------------------------
'''''Enter the transaction code
Call SetTcode(DT_OKCD_OCC1) 
Call PressEnter()     ' 
Call TakeScreenShot

Call ClickButton("Get variant\.\.\.   \(Shift\+F5\)",False)
Call TakeScreenShot
Call SetTextbox("Created By","ENAME-LOW","","",False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",True)
Wait 2
Call FindRowNumber("Variant Catalog for Program RWVKP007.*", "Variant Name", "AB_AB MAINBAN", "DT_ROW_NUMBER")
Call SelectRowGuiGridbyRowNo("Variant Catalog for Program RWVKP007.*", 0, DT_ROW_NUMBER, True)
Call takeScreenShot()
Wait 2
Call ClickButtonIfExist("Choose   \(F2\)",True)
Wait 5
Call takeScreenShot()

Call SetTextbox("Article","S_MATNR-LOW","",DT_NEW_ARTICLE,False)
Call SetTextbox("to","S_DATUM-HIGH","",ConvertDate(DT_VALIDITY_TO),False)
Call takeScreenShot()
Call ClickButton("Execute   \(F8\)",FalsE)
Call takeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call takeScreenShot()
Call GetStatusBar("item1","DT_ZMDAM_BOM_DATA_ENTRY_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call  VerifyStatusBar("Data saved; pricing document "&DT_ZMDAM_BOM_DATA_ENTRY_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" created")


Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call PressEnter()

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************







