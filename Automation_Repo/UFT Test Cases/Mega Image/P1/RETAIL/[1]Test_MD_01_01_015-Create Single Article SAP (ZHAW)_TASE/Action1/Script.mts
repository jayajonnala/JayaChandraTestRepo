
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_01_01_015-Create Single Article SAP (ZHAW)
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

gstrTestCaseName = "Test_MD_01_01_015-Create Single Article SAP (ZHAW)"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\FICO\TASE_DT_AA001 - Scrapping Asset partial Deduct with NBV, IFRS equal Loca.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'----------------------Tcode MM41 ----------------------------
'Enter the transaction code
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
Call TakeScreenShot()
Call SetComboByKey("RMMW1-MTART",DT_MM41_0100_ARTICLE_TYPE)
Call SetTextbox("Mdse Catgry","RMMW1-MATKL","",DT_MM41_0100_MDSE_CATGRY,False)
Call SetComboByKey("RMMW1-ATTYP",DT_MM41_0100_ARTL_CATEGORY)
Call SetTextbox("Purchasing Org\.","RMMW1-EKORG","",DT_MM41_0100_PURCHASING_ORG,False)
''''Call SetTextbox("Vendor","RMMW1-LIFNR","",DT_MM41_0100_VENDOR,False)
Call SetTextboxNoLabel("RMMW1-LIFNR","",DT_MM41_0100_VENDOR,False)
Call TakeScreenShot()
Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100","Screen description","Basic Data",False)
Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100","Screen description","Purchasing",False)
Call TakeScreenShot()
Call PressEnter() 

Call SetTableData("SAPLMGD2TC_ME_8022", "AUoM", 1, "", "", DT_MM41_8022_TABLECELL_AUOM_0, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "AUoM", 2, "", "", DT_MM41_8022_TABLECELL_AUOM_1, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "AUoM", 3, "", "", DT_MM41_8022_TABLECELL_AUOM_2, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "AUoM", 4, "", "", DT_MM41_8022_TABLECELL_AUOM_3, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Number", 2, "", "", DT_MM41_8022_TABLECELL_NUMBER_1, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Number", 3, "", "", DT_MM41_8022_TABLECELL_NUMBER_2, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Number", 4, "", "", DT_MM41_8022_TABLECELL_NUMBER_3, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "OUn", 2, "", "", "ON", False)
Call SetTableData("SAPLMGD2TC_ME_8022", "D/I", 2, "", "", "ON", False)
Call SetTableData("SAPLMGD2TC_ME_8022", "SUn", 1, "", "", "ON", False)
Call SetTableData("SAPLMGD2TC_ME_8022", "LUn", 2, "", "",DT_MM41_8022_TABLECELL_LUN_1, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "LUn", 3, "", "",DT_MM41_8022_TABLECELL_LUN_2, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "LUn", 4, "", "",DT_MM41_8022_TABLECELL_LUN_3, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "GTIN", 1, "", "",DT_MM41_8022_TABLECELL_GTIN_0, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Gross Weight", 1, "", "",DT_MM41_8022_TABLECELL_GROSS_WEIGHT_0, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Gross Weight", 2, "", "",DT_MM41_8022_TABLECELL_GROSS_WEIGHT_1, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Gross Weight", 3, "", "",DT_MM41_8022_TABLECELL_GROSS_WEIGHT_2, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Gross Weight", 4 ,"", "",DT_MM41_8022_TABLECELL_GROSS_WEIGHT_3, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Net Weight", 1, "", "",DT_MM41_8022_TABLECELL_NET_WEIGHT_0, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Wt", 2, "", "",DT_MM41_8022_TABLECELL_WT_1, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Wt", 3, "", "",DT_MM41_8022_TABLECELL_WT_2, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Wt", 4, "", "",DT_MM41_8022_TABLECELL_WT_3, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Length", 1, "", "",DT_MM41_8022_TABLECELL_LENGTH_0, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Length", 2, "", "",DT_MM41_8022_TABLECELL_LENGTH_1, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Length", 3, "", "",DT_MM41_8022_TABLECELL_LENGTH_2, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Length", 4 ,"", "",DT_MM41_8022_TABLECELL_LENGTH_3, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Width", 1, "", "",DT_MM41_8022_TABLECELL_WIDTH_0, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Width", 2, "", "",DT_MM41_8022_TABLECELL_WIDTH_1, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Width", 3, "", "",DT_MM41_8022_TABLECELL_WIDTH_2, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Width", 4 ,"", "",DT_MM41_8022_TABLECELL_WIDTH_3, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Height", 1, "", "",DT_MM41_8022_TABLECELL_HEIGHT_0, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Height", 2, "", "",DT_MM41_8022_TABLECELL_HEIGHT_1, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Height", 3, "", "",DT_MM41_8022_TABLECELL_HEIGHT_2, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Height", 4 ,"", "",DT_MM41_8022_TABLECELL_HEIGHT_3, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Unit of Dimension", 1, "", "",DT_MM41_8022_TABLECELL_UNIT_OF_DIMENSION_0, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Unit of Dimension", 2, "", "",DT_MM41_8022_TABLECELL_UNIT_OF_DIMENSION_1, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Unit of Dimension", 3, "", "",DT_MM41_8022_TABLECELL_UNIT_OF_DIMENSION_2, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Unit of Dimension", 4 ,"", "",DT_MM41_8022_TABLECELL_UNIT_OF_DIMENSION_3, False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 

Call SetTextbox("ABC Indicator","MAW1-WMAAB","",DT_MM41_2001_ABC_INDICATOR,False)
Call SetTextbox("Prod\. Hierarchy","MARA-PRDHA",""," ",False)
Call SetTextbox("Division","MARA-SPART","",DT_MM41_2001_DIVISION,False)

'Capture the screenshot
Call TakeScreenShot()

'Click on Tax Data
Call ClickButton("More Tax Data",False)
Wait(2)

Call SetTableDataNoRef("SAPLMGD2TC_STEUERN","#1",2,DT_MM41_2181_TABLECELL__1,False)
Call SetTableDataNoRef("SAPLMGD2TC_STEUERN","#3",2,DT_MM41_2181_TABLECELL__1_OCC1,False)
Call SetTableDataNoRef("SAPLMGD2TC_STEUERN","#5",2,DT_MM41_2181_TABLECELL__1_OCC2,False)
Call TakeScreenShot()

''Click on Go to main data
Call ClickButton("Go to main data   \(Ctrl\+Shift\+F3\)",False)
Wait(2)

Call SetTextbox("Ctry of origin","MAW1-WHERL","",DT_MM41_2002_CTRY_OF_ORIGIN,False)
Call SetTextbox("Temperature","MARA-TEMPB","",DT_MM41_2002_TEMPERATURE,False)
Call SetTextbox("Stor\.conditions","MARA-RAUBE","",DT_MM41_2002_STORCONDITIONS,False)
Call TakeScreenShot()
Call SetTextbox("Purch\. Group","MAW1-WEKGR","",DT_MM41_2003_PURCH_GROUP,False)
Call SetTextbox("Supply source","MARA-BWSCL","",DT_MM41_2003_SUPPLY_SOURCE,False)
Call SetTextbox("Trans\. Group","MARA-TRAGR","",DT_MM41_2003_TRANS_GROUP,False)
Call SetTextbox("Loading Group","MAW1-WLADG","",DT_MM41_2003_LOADING_GROUP,False)
Call TakeScreenShot()

''Click on Next Page
Call ClickButton("Next Page",False)
Wait(2)

''Click on Next Page
Call ClickButton("Next Page",False)
Wait(2)
Call TakeScreenShot()
Call SetTextbox("Country of Origin","RCTMS-MWERT","",DT_MM41_4000_RCTMSMWERT,False)
Call SetTextbox("Item strategy","RCTMS-MWERT","",DT_MM41_4000_RCTMSMWERT_OCC1,False)
Call SetTextbox("Multi-tier assortment pricing","RCTMS-MWERT","",DT_MM41_4000_RCTMSMWERT_OCC2,False)
Call SetTextbox("Price strategy Maxi","RCTMS-MWERT","",DT_MM41_4000_RCTMSMWERT_OCC3,False)
Call SetTextbox("Price strategy Tempo","RCTMS-MWERT","",DT_MM41_4000_RCTMSMWERT_OCC4,False)

'''Call SendKey("F4")
''' SetTableData(tableName, columnName, rowNumber, refColumnName, refCellValue, cellValue, blnIsItPopup)
''' SelectCellGuiTable(tableName, columnName, refColumnName, refColumnVal, blnIsItPopup)
'''Call SelectCellGuiTable("SAPLCTMSVALUE_S","Field for Selecting an Entry",2,"","","ON",False)
''
''' ClickCellTableByRowNo(tableName, columnName, rowNumber, blnIsItPopup)
''' ClickCellTable(tableName, columnName, rowNumber, columnNameRef, tableValRef, blnIsItPopup)
'''Call Clickbutton("Continue   \(F8\)",False)
'''Call TakeScreenShot()

'Click on Next Page
Call ClickButton("Next Page",False)
Wait(2)
Call ClickButton("Next Page",False)
Call TakeScreenShot()
Call SetTextbox("Recommended glass","RCTMS-MWERT","",DT_MM41_4000_RCTMSMWERT_OCC5,False)

''Click on Next Page
Call ClickButton("Next Page",False)
Wait(2)
Call TakeScreenShot()

Call SetTableData("SAPLWRF_ARTICLE_SCREENSTC_AH_ASSIGNM","Main Assignment",1,"","",DT_MM41_2021_TABLECELL_MAIN_ASSIGNMENT_0,False)
Call SetTableData("SAPLWRF_ARTICLE_SCREENSTC_AH_ASSIGNM","Hierarchy Node",1,"","",DT_MM41_2021_TABLECELL_HIERARCHY_NODE_0,False)
Call SetTextBox("Brand","MARA-ZZBRAND","",DT_MM41_0001_BRAND,False)
Call SetTextBox("Sub Brand","MARA-ZZSUB_BRAND","",DT_MM41_0001_SUB_BRAND,False)
Call TakeScreenShot()
Call ClickButton("Display item details",False)
If SAPGuiSession(sessionObject).SapGuiWindow(windowobject).SAPGuiEdit("guicomponenttype:=32","attachedtext:=Unit weight \(gr\)","name:=ZMDAM_MARAEXT-PRIVATE_BRAND","Index:=0").Exist(5)=False Then
Call ClickButton("Display item details",False)
End If
' SelectCheckbox(checkboxName, checkBoxIndex, OnOffStatus, blnIsItPopup)
Call TakeScreenShot()
Call SelectTab("INFOEXT","Sales",False)
Call TakeScreenShot()
Call SelectCheckbox("ZMDAM_MARAEXT-FAIR_TRADE",0,"ON",False)
Call SelectCheckbox("ZMDAM_MARAEXT-PRESENCE_OF_BBD",0,"ON",False)
Call SetTextBox("Unit weight \(gr\)","ZMDAM_MARAEXT-UNIT_WEIGHT","",DT_MM41_4007_UNIT_WEIGHT__GR,False)
Call TakeScreenShot()
''Call SetTextBox("Unit weight \(gr\)","ZMDAM_MARAEXT-PRIVATE_BRAND","",DT_MM41_4007_UNIT_WEIGHT__GR_OCC1,False)
Call TakeScreenShot()
Call ClickCellTableByRowNo("SAPLZMDAM_EXT_SCRTC_LTEXT","#1",1,False)
Call TakeScreenShot()
Call ClickCellTableByRowNo("SAPLZMDAM_EXT_SCRTC_LONGTEXT","#1",1,False)
Call TakeScreenShot()
Call ClickButton("Create text",False)
 ''SetComboByKey(attachedTextOrComboName, keyValue)
 Call TakeScreenShot()
Call SetComboByKey("Create text in",DT_MM41_0005_CREATE_TEXT_IN)
Call PressEnter()
Call TakeScreenShot()
' SetTextArea(strTextAreaVal)
Call SetTextArea("Test automation%cr%")
Call ClickButton("Create text",False)
Call TakeScreenShot()
Call SetComboByKey("Create text in",DT_MM41_0005_CREATE_TEXT_IN_OCC1)
' SetComboByKey(attachedTextOrComboName, keyValue)
Call SetComboByKey("Copy from",DT_MM41_0005_COPY_FROM)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot()
Call SelectTab("TABSPR1"," Purchasing",False)
 'SelectTab(tabStripName, tabName, blnIsItPopup)
Call TakeScreenShot()
 'SelectCheckbox(checkboxName, checkBoxIndex, OnOffStatus, blnIsItPopup)
Call SelectCheckbox("EINA-RELIF",0,"ON",False)
Call SetTextBox("Available from","EINA-LIFAB","",ConvertDate(DT_MM41_2221_AVAILABLE_FROM),False)
Call SetTextBox("Available to","EINA-LIFBI","",DT_MM41_2221_AVAILABLE_TO,False)
'''Call SetTextBox("Vendor artl no\.","EINA-IDNLF","",DT_MM41_2221_VENDOR_ARTL_NO,False)
Call SetTextBox("Vendor artl no\.","EINA-IDNLF","",DT_MM41_2221_VENDOR_ARTL_NO,False)
Call SetTextBox("Net Price","EINE-NETPR","",DT_MM41_2223_NET_PRICE,False)
Call PressEnter()
Call TakeScreenShot()
Call ClickButton("Cancel",True)
Call ClickButton("For the Conditions",False)
 'VerifyTableCellContent(tableRowNumber, tableColumnName, tableName, expectedValue)
Call VerifyTableCellContent(1,"CnTy","SAPMV13ATCTRL_D0201",DT_MM41_0201_CHECK_TEXT_OF_TABLECELL_CNTY_0)
Call VerifyTablecellContent(1,"Name","SAPMV13ATCTRL_D0201",LCase(DT_MM41_0201_CHECK_TEXT_OF_TABLECELL_NAME_0))
Call TakeScreenShot()
Call ClickButton("Back   \(F3\)",False)
Call SetTextBox("Minimum Qty","EINE-MINBM","",DT_MM41_2222_MINIMUM_QTY,False)
Call TakeScreenShot()
Call SetTextBox("Var\. Order Unit","EINA-VABME","",DT_MM41_2221_VAR_ORDER_UNIT,False)
Call ClickButton("Go to additional data   \(Ctrl\+F6\)",False)
Call TakeScreenShot()
' SelectCellGuiTable(tableName, columnName, refColumnName, refColumnVal, blnIsItPopup)
Call SelectCellGuiTable("SAPLMGD2TC_KTXT","Language","RO","",False)
Call ClickButton("Delete Row",False)
Call ClickButton("Delete Row",False)
Call ClickButton("Delete Row",False)
Call ClickButton("Delete Row",False)
' SetTableData(tableName, columnName, rowNumber, refColumnName, refCellValue, cellValue, blnIsItPopup)
Call SetTableData("SAPLMGD2TC_KTXT", "Material Description", 1, "", "", DT_MM41_8000_TABLECELL_ARTICLE_DESCRIPTION_0, False)
Call SetTableData("SAPLMGD2TC_KTXT", "Material Description", 2, "", "", DT_MM41_8000_TABLECELL_ARTICLE_DESCRIPTION_1, False)
Call TakeScreenShot()
Call SelectTab("TABSPR1","Additional GTINs",False)
Call TakeScreenShot()
Call SetTableData("SAPLMGD2TC_EAN", "AV", 2, "", "", "ON", False)
Call SetTableData("SAPLMGD2TC_EAN", "MV", 2, "", "", "ON", False)
' SetTableData(tableName, columnName, rowNumber, refColumnName, refCellValue, cellValue, blnIsItPopup)
Call TakeScreenShot()
Call ClickButton("Go to main data   \(Ctrl\+Shift\+F3\)",False)
Call SelectTab("TABSPR1","POS",False)
Call TakeScreenShot()
Call SetTableData("SAPLMGD2TC_BON","L",1,"","",DT_MM41_2273_TABLECELL_L_0,False)
Call SetTableData("SAPLMGD2TC_BON","L",2,"","",DT_MM41_2273_TABLECELL_L_1,False)
Call SetTableData("SAPLMGD2TC_BON","AUn",1,"","",DT_MM41_2273_TABLECELL_AUN_0,False)
Call SetTableData("SAPLMGD2TC_BON","AUn",2,"","",DT_MM41_2273_TABLECELL_AUN_1,False)
Call SetTableData("SAPLMGD2TC_BON","Till rcpt texts for unit",1,"","",DT_MM41_2273_TABLECELL_TILL_RCPT_TEXTS_FOR_UNIT_0,False)
Call SetTableData("SAPLMGD2TC_BON","Till rcpt texts for unit",2,"","",DT_MM41_2273_TABLECELL_TILL_RCPT_TEXTS_FOR_UNIT_1,False)
Call TakeScreenShot()
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot()
' GetStatusBar(itemNo, dataTableColumnName)
Call GetStatusBar("item1","DT_MM41_0100_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_MM41_0100_CHECK_TEXT_OF_STATUSBAR_OUTPUT",DT_MM41_0100_CHECK_TEXT_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
wait 5
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'----------------------Tcode WSM3 ----------------------------
'Enter the transaction code
Call SetTcode(DT_WSM3) 
Call PressEnter()     ' 

Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextBox("Variant","V-LOW","",DT_VARIANT,True)
Call SetTextBox("Created By","ENAME-LOW","",DT_CREATED_BY,True)
Call ClickButton("Execute   \(F8\)",False)
Call SetTextBox("Assortment","ASORT-LOW",1,DT_ASSORTMENT,False)
Call TakeScreenShot()
Call SetTextBox("Article","MATNR-LOW",2,DT_MM41_0100_CHECK_TEXT_OF_STATUSBAR,False)
Call SetTextBox("Listing Valid From","DATAB",4,ConvertDate(DT_VALID_FROM),False)
Call ClickButton("Execute   \(F8\)",False)
' VerifyStatusBar(Content)
Call VerifyStatusBar(DT_STATUS_BAR)

'''----------------------Tcode ZMDAS_WSL11 ----------------------------
''Enter the transaction code
Call SetTcode(DT_MM41_0100_OKCD) 
Call PressEnter()     ' 

Call SetTextBox("Assortment","S_FILIA-LOW",1,DT_MM41_1000_ASSORTMENT,False)
Call SetTextBox("Article","S_ARTNR-LOW",2,DT_MM41_1000_ARTICLE,False)
Call SetTextBox("Customer No\. - Site","S_LOCNR-LOW",5,"R*",False)
Call TakeScreenShot()
Call SetTextBox("Valid From","P_DATAB",3,ConvertDate(DT_MM41_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DATAB),False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
 'VerifyGridCellContentbyName(gridName, gridRowNumber, gridColumnName, gridIndex, expectedValue)
Call VerifyGridCellContentbyName("shell",1,"Article","",DT_MM41_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ARTNR)
Call VerifyGridCellContentbyName("shell",2,"Customer No. - Site","",DT_MM41_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOCNR)
Call VerifyGridCellContentbyName("shell",2,"Assortment","",DT_MM41_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_FILIA)
Call TakeScreenShot()
'
'''''''----------------------Tcode MM42 ----------------------------
''''''''Enter the transaction code
''''''''''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''''
'''Call SetTcode(DT_MM41_0500_OKCD) 
'''Call PressEnter()     ' 
'''
'''Call SetTextBox("Article","RMMW1-MATNR","",DT_MM41_0100_ARTICLE,False)
'''Wait(5)
'''Call SetTextBox("Purchasing Org\.","RMMW1-EKORG","","",False)
'''Call SetTextBoxNoLabel("RMMW1-LIFNR","","",False)
'''Call SetTextBox("Sales Org\.","RMMW1-VKORG","",DT_MM41_0100_SALES_ORG,False)
'''Call SetTextBox("Distr\. Channel","RMMW1-VTWEG","",DT_MM41_0100_DISTR_CHANNEL,False)
'''Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100","Screen description","Basic Data",False)
'''Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100","Screen description","Sales",False)
'''Call TakeScreenShot()
'''Call PressEnter()
''''' GTIN step added to avoid the errors in posting
'''Call SetTableData("SAPLMGD2TC_ME_8022", "GTIN", 1, "AUoM", "EA", DT_GTIN_NUM, False)
'''Call SelectTab("TABSPR1"," Sales",False)
'''Call SetTextBox("Sales Unit","MVKE-VRKME","","",False)
'''Call TakeScreenShot()
'''Call ClickButton("Other Sales Data",False)
'''Call SetTextBox("Article pricing grp","MVKE-KONDM","",DT_MM41_2152_ARTICLE_PRICING_GRP,False)
'''Call TakeScreenShot()
'''Call ClickButton("Go to main data   \(Ctrl\+Shift\+F3\)",False)
'''Call ClickButton("Switch area of validity   \(Shift\+F1\)",False)
'''Call SetTextBox("Sales Unit","RMMW1-VRKME","",DT_MM41_0081_SALES_UNIT,True)
'''Call TakeScreenShot()
'''Call SetTextBox("Sales Org\.","RMMW1-VKORG","",DT_MM41_0081_SALES_ORG,True)
'''Call SetTextBox("Distr\. Channel","RMMW1-VTWEG","",DT_MM41_0081_DISTR_CHANNEL,True)
'''Call PressEnter()
'''Call SetTextBox("Final Price","CALP-ENDPR","",DT_MM41_2233_FINAL_PRICE,False)
'''Call TakeScreenShot()
'''Call PressEnter()
'''Call VerifyStatusBar(DT_MM41_4030_CHECK_TEXT_OF_STATUSBAR)
''''Call GetStatusBar("item2","DT_MM41_4030_CHECK_TEXT_OF_STATUSBAR")
'''Call ClickButton("Save   \(Ctrl\+S\)",False)
''''Call GetStatusBar("item2","DT_MM41_0100_CHECK_TEXT_OF_STATUSBAR_OCC2")
'''Call VerifyStatusBar(DT_MM41_0100_CHECK_TEXT_OF_STATUSBAR_OCC2)
'''Call TakeScreenShot()

''----------------------Tcode VKP5 - Setting sales price steps is changed from MM42 to VKP5 ----------------------------
Call SetTcode(DT_SET_TCODE_VKP5) 
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot

Call ClickButton("btn\[17\]",False)
Call SetTextbox("Variant","V-LOW","",DT_VKP5_VARIANT,True)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",True)
Call TakeScreenShot
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Article","S_MATNR-LOW","",DT_MM41_0100_ARTICLE,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)

Call TakeScreenShot
Call SetGridData("", 1, "ENDPR", DT_MM41_2233_FINAL_PRICE, False)
Call PressEnter()     
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call GetStatusBar("MessageType", "DT_VKP5_1000_GET_MESSAGE_TYPE_OUTPUT")
Call GetStatusBar("item1", "DT_ZMDAM_BOM_REPORT_1000_GET_PRICING_DOCUMENT_OUTPUT")
Call VerifyStatusBar("Data saved; pricing document "&DT_ZMDAM_BOM_REPORT_1000_GET_PRICING_DOCUMENT_OUTPUT&" created")
Call TakeScreenShot

'''----------------------Tcode VKP2 ----------------------------
'Enter the transaction code
Call SetTcode(DT_MM41_0100_OKCD_OCC1) 
Call PressEnter() 

Call SetTextBox("Article","S_MATNR-LOW",1,DT_MM41_1000_ARTICLE_OCC1,False)
Call SetTextBox("Sales Organization","S_VKORG-LOW",2,DT_MM41_1000_SALES_ORGANIZATION,False)
Call SetTextBox("Distribution Channel","S_VTWEG-LOW",3,DT_MM41_1000_DISTRIBUTION_CHANNEL,False)
Call SetTextBox("Validity Period","S_DATUM-LOW","",ConvertDate(DT_MM41_1000_VALIDITY_PERIOD),False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
' VerifyGridCellContent(gridTitle, gridRowNumber, gridColumnName, gridIndex, expectedValue)
Call VerifyGridCellContent("Sales Price Conditions",1,"Article","",DT_MM41_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR)
Call VerifyGridCellContent("Sales Price Conditions",1,"Sales unit","",DT_MM41_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VRKME)
Call VerifyGridCellContent("Sales Price Conditions",1,"Sales Organization","",DT_MM41_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VKORG)
Call VerifyGridCellContent("Sales Price Conditions",1,"Valid from","",ConvertDate(DT_MM41_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DATAB_OCC1))
Call VerifyGridCellContent("Sales Price Conditions",1,"Condition rate","",DT_MM41_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KBETR)
Call VerifyGridCellContent("Sales Price Conditions",1,"Condition Type","",DT_MM41_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KSCHL)
Call TakeScreenShot()

Call LogOff()
Call FinalStatus ()

''''Enter the details
'''Call SetTextbox("Article","RMMW1-MATNR","",DT_MM42_0100_ARTICLE,False)
'''Call SetTextbox("Sales Org\.","RMMW1-VKORG","",DT_MM42_0100_SALES_ORG,False)
'''Call SetTextbox("Distr\. Channel","RMMW1-VTWEG","",DT_MM42_0100_DISTR_CHANNEL,False)
'''
''''Capture the screenshot
'''Call TakeScreenShot()
'''
'''Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100","Screen description","Basic Data",False)
'''Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100","Screen description","Purchasing",False)
''''Capture the screenshot
'''Call TakeScreenShot()
'''Call PressEnter()
'''wait(2)
'''Call TakeScreenShot()
'''Call PressEnter()
'''wait(2)
'''
''''Navigate to Purchasing Tab
'''Call SelectTab("TABSPR1"," Purchasing",False)
'''Wait(1)
'''Call TakeScreenShot()
'''
'''Call SetTextbox("Rem\. Shelf Life","EINE-MHDRZ","",DT_MM42_2222_REM_SHELF_LIFE,False)
'''Call SetTextbox("Minimum Qty","EINE-MINBM","",DT_MM42_2222_MINIMUM_QTY,False)
'''Call SetTextbox("Pl\. Deliv\. Time","EINE-APLFZ","",DT_MM42_2222_PL_DELIV_TIME,False)
''''Capture the screenshot
'''Call TakeScreenShot()
'''
''''Save the Details
''''Call ClickButton("Save   \(Ctrl\+S\)",False)
''''Wait(2)
''''Call TakeScreenShot()
''''
'''''Click on Enterif pop up appears
''''Call ClickButtonIfExist("Enter   \(F5\)",True)
''''Wait(2)
''''
'''''Verify the Message Type
''''Call VerifyStatusBarMessageType("S")
''''
'''''Verify the status bar message
''''Call VerifyStatusBar(DT_MM42_0100_CHECK_TEXT_OF_STATUSBAR)
'''''
'''
'------------------------'Log Off From Applicaton--------------------------------



'*********************************************End Of Script*********************************************************************









'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [8,3437477]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//




