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
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'reload DS to update dates and calculations
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_MD_01_01_015-Create Single Article SAP (ZFER)_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 17th June
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_MD_01_01_015-Create Single (ZFER)_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_MD_01_01_015-Create Single Article SAP (ZFER)_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'''''Increment the parameter/reload
''''Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
''''Call WriteRunTimeDataToExcelGlobalSheet ("DT_XYZ",1)
''''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'''''
''''''''----------------------Tcode MM41----------------------------
'''''Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
''Capture the screenshot
Call TakeScreenShot()

Call SetComboByKey("RMMW1-MTART",DT_MM41_0100_ARTICLE_TYPE)
Call SetTextbox("Mdse Catgry","RMMW1-MATKL","",DT_MM41_0100_MDSE_CATGRY,false)
Call SetComboByKey("Artl category",DT_MM41_0100_ARTL_CATEGORY)
Call SetTextbox("Purchasing Org.","RMMW1-EKORG","",DT_MM41_0100_PURCHASING_ORG,false)
'''Call SetTextbox("Vendor","RMMW1-LIFNR","",DT_MM41_0100_VENDOR,false)
Call SetTextboxNoLabel("RMMW1-LIFNR","",DT_MM41_0100_VENDOR,false)
Call SelectRowGuiTableByRow("SAPLMGMWTAB_CONT_0100",3,False)

Call PressEnter() 
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot

Call SetTableData("SAPLMGD2TC_ME_8022","SUn",1,"<NA>","<NA>",DT_MM41_8022_TABLECELL_SUN_0,False)
Call SetTableData("SAPLMGD2TC_ME_8022","OUn",2,"<NA>","<NA>",DT_MM41_8022_TABLECELL_OUN_1,False)
Call SetTableData("SAPLMGD2TC_ME_8022","D/I",2,"<NA>","<NA>",DT_MM41_8022_TABLECELL_DI_1,False)

Call SetTableData("SAPLMGD2TC_ME_8022","AUoM",1,"<NA>","<NA>",DT_MM41_8022_TABLECELL_AUOM_0,False)
Call SetTableData("SAPLMGD2TC_ME_8022","AUoM",2,"<NA>","<NA>",DT_MM41_8022_TABLECELL_AUOM_1,False)
Call SetTableData("SAPLMGD2TC_ME_8022","AUoM",3,"<NA>","<NA>",DT_MM41_8022_TABLECELL_AUOM_2,False)
Call SetTableData("SAPLMGD2TC_ME_8022","AUoM",4,"<NA>","<NA>",DT_MM41_8022_TABLECELL_AUOM_3,False)

Call SetTableData("SAPLMGD2TC_ME_8022","Number",2,"<NA>","<NA>",DT_MM41_8022_TABLECELL_NUMBER_1,False)
Call SetTableData("SAPLMGD2TC_ME_8022","Number",3,"<NA>","<NA>",DT_MM41_8022_TABLECELL_NUMBER_2,False)
Call SetTableData("SAPLMGD2TC_ME_8022","Number",4,"<NA>","<NA>",DT_MM41_8022_TABLECELL_NUMBER_3,False)

Call SetTableData("SAPLMGD2TC_ME_8022","LUn",2,"<NA>","<NA>",DT_MM41_8022_TABLECELL_LUN_1,False)
Call SetTableData("SAPLMGD2TC_ME_8022","LUn",3,"<NA>","<NA>",DT_MM41_8022_TABLECELL_LUN_2,False)
Call SetTableData("SAPLMGD2TC_ME_8022","LUn",4,"<NA>","<NA>",DT_MM41_8022_TABLECELL_LUN_3,False)

Call SetTableData("SAPLMGD2TC_ME_8022","GTIN",1,"<NA>","<NA>",DT_MM41_8022_TABLECELL_GTIN_0_O,False)

Call SetTableData("SAPLMGD2TC_ME_8022","Gross Weight",1,"<NA>","<NA>",DT_MM41_8022_TABLECELL_GROSS_WEIGHT_0,False)
Call SetTableData("SAPLMGD2TC_ME_8022","Gross Weight",2,"<NA>","<NA>",DT_MM41_8022_TABLECELL_GROSS_WEIGHT_1,False)
Call SetTableData("SAPLMGD2TC_ME_8022","Gross Weight",3,"<NA>","<NA>",DT_MM41_8022_TABLECELL_GROSS_WEIGHT_2,False)
Call SetTableData("SAPLMGD2TC_ME_8022","Gross Weight",4,"<NA>","<NA>",DT_MM41_8022_TABLECELL_GROSS_WEIGHT_3,False)

Call SetTableData("SAPLMGD2TC_ME_8022","Net Weight",1,"<NA>","<NA>",DT_MM41_8022_TABLECELL_NET_WEIGHT_0,False)

Call SetTableData("SAPLMGD2TC_ME_8022","Wt",2,"<NA>","<NA>",DT_MM41_8022_TABLECELL_WT_1,False)
Call SetTableData("SAPLMGD2TC_ME_8022","Wt",3,"<NA>","<NA>",DT_MM41_8022_TABLECELL_WT_2,False)
Call SetTableData("SAPLMGD2TC_ME_8022","Wt",4,"<NA>","<NA>",DT_MM41_8022_TABLECELL_WT_3,False)

Call SetTableData("SAPLMGD2TC_ME_8022","Length",1,"<NA>","<NA>",DT_MM41_8022_TABLECELL_LENGTH_0,False)
Call SetTableData("SAPLMGD2TC_ME_8022","Length",2,"<NA>","<NA>",DT_MM41_8022_TABLECELL_LENGTH_1,False)
Call SetTableData("SAPLMGD2TC_ME_8022","Length",3,"<NA>","<NA>",DT_MM41_8022_TABLECELL_LENGTH_2,False)
Call SetTableData("SAPLMGD2TC_ME_8022","Length",4,"<NA>","<NA>",DT_MM41_8022_TABLECELL_LENGTH_3,False)

Call SetTableData("SAPLMGD2TC_ME_8022","Width",1,"<NA>","<NA>",DT_MM41_8022_TABLECELL_WIDTH_0,False)
Call SetTableData("SAPLMGD2TC_ME_8022","Width",2,"<NA>","<NA>",DT_MM41_8022_TABLECELL_WIDTH_1,False)
Call SetTableData("SAPLMGD2TC_ME_8022","Width",3,"<NA>","<NA>",DT_MM41_8022_TABLECELL_WIDTH_2,False)
Call SetTableData("SAPLMGD2TC_ME_8022","Width",4,"<NA>","<NA>",DT_MM41_8022_TABLECELL_WIDTH_3,False)

Call SetTableData("SAPLMGD2TC_ME_8022","Height",1,"<NA>","<NA>",DT_MM41_8022_TABLECELL_HEIGHT_0,False)
Call SetTableData("SAPLMGD2TC_ME_8022","Height",2,"<NA>","<NA>",DT_MM41_8022_TABLECELL_HEIGHT_1,False)
Call SetTableData("SAPLMGD2TC_ME_8022","Height",3,"<NA>","<NA>",DT_MM41_8022_TABLECELL_HEIGHT_2,False)
Call SetTableData("SAPLMGD2TC_ME_8022","Height",4,"<NA>","<NA>",DT_MM41_8022_TABLECELL_HEIGHT_3,False)

Call SetTableData("SAPLMGD2TC_ME_8022","Unit of Dimension",1,"<NA>","<NA>",DT_MM41_8022_TABLECELL_UNIT_OF_DIMENSION_0,False)
Call SetTableData("SAPLMGD2TC_ME_8022","Unit of Dimension",2,"<NA>","<NA>",DT_MM41_8022_TABLECELL_UNIT_OF_DIMENSION_1,False)
Call SetTableData("SAPLMGD2TC_ME_8022","Unit of Dimension",3,"<NA>","<NA>",DT_MM41_8022_TABLECELL_UNIT_OF_DIMENSION_2,False)
Call SetTableData("SAPLMGD2TC_ME_8022","Unit of Dimension",4,"<NA>","<NA>",DT_MM41_8022_TABLECELL_UNIT_OF_DIMENSION_3,False)
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("ABC Indicator","MAW1-WMAAB","",DT_MM41_2001_ABC_INDICATOR,false)
Call SetTextbox("Prod\. Hierarchy","MARA-PRDHA","","",false)
Call SetTextbox("Division","MARA-SPART","",DT_MM41_2001_DIVISION,false)
Call TakeScreenShot
Call ClickButton("MLAN_STEU",False)  ''Tax data button

Call SetTableData("SAPLMGD2TC_STEUERN","#1",2,"<NA>","<NA>",DT_MM41_2181_TABLECELL__1,False)
Call SetTableData("SAPLMGD2TC_STEUERN","#3",2,"<NA>","<NA>",DT_MM41_2181_TABLECELL__1_OCC1,False)
Call SetTableData("SAPLMGD2TC_STEUERN","#5",2,"<NA>","<NA>",DT_MM41_2181_TABLECELL__1_OCC2,False)
Call TakeScreenShot
Call ClickButton("Go to main data   \(Ctrl\+Shift\+F3\)",False)

Call SetTextbox("Ctry of origin","MAW1-WHERL","",DT_MM41_2002_CTRY_OF_ORIGIN,false)
Call SetTextbox("Stor.conditions","MARA-RAUBE","",DT_MM41_2002_STORCONDITIONS,false)
Call SetTextbox("Temperature","MARA-TEMPB","",DT_MM41_2002_TEMPERATURE,false)
Call TakeScreenShot

Call SetTextbox("Purch. group","MAW1-WEKGR","",DT_MM41_2003_PURCH_GROUP,false)
Call SetTextbox("Supply source","MARA-BWSCL","",DT_MM41_2003_SUPPLY_SOURCE,false)
Call SetTextbox("Trans. Group","MARA-TRAGR","",DT_MM41_2003_TRANS_GROUP,false)
Call SetTextbox("Loading Group","MAW1-WLADG","",DT_MM41_2003_LOADING_GROUP,false)
Call TakeScreenShot

'Call SendKey("{PGDN}")
'Wait(1)
'Call SendKey("{PGDN}")
'Wait(1)
'Call SendKey("{PGDN}")
'Wait(1)

Call ClickButton("Next Page",False)  
Call ClickButton("Next Page",False)  
Call ClickButton("Next Page",False)  
Call ClickButton("First Page",False)  
Call ClickButton("Last Page",False)  
Call ClickButton("Previous Page",False)  

'Call FocusTextBox("IF BIO","RCTMS-MWERT",False)
'Wait(2)
'Call ClickButton("OES_PLAST",False)  
'Wait(2)
'Call ClickButton("OES_PUP",False)  

Call SetTextbox("Electronic shelf label","RCTMS-MWERT","",DT_MM41_4000_RCTMSMWERT,false)
Call SetTextbox("Recommended glass","RCTMS-MWERT","",DT_MM41_4000_RCTMSMWERT_OCC1,false)
Call SetTextbox("Color","RCTMS-MWERT","",DT_MM41_4000_RCTMSMWERT_OCC2,false)
Call SetTextbox("Package type","RCTMS-MWERT","",DT_MM41_4000_RCTMSMWERT_OCC3,false)
Call SetTextbox("Sales relevant indicator APPRO","RCTMS-MWERT","",DT_MM41_4000_RCTMSMWERT_OCC4,false)
Call TakeScreenShot

Call FocusTextBox("Sales relevant indicator APPRO","RCTMS-MWERT",False)
Wait(2)
Call SendKey("{F4}")

'Set objWsh = CreateObject("WScript.Shell") 
'objWsh.SendKeys "{F4}"
'Set objWsh=nothing

Wait(2)
Call SetTableDataNoRef("SAPLCTMSVALUE_S","Field for Selecting an Entry",2,DT_MM41_1900_TABLECELL_FIELD_FOR_SELECTING_AN_ENTRY_1,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(F8\)",True)

''Call ClickButton("OES_PDOWN",False)
Call ClickButton("Next Page",False)
Wait(2)
Call SetTextbox("Bio or Eco","RCTMS-MWERT","",DT_MM41_4000_RCTMSMWERT_OCC5,false)
Call TakeScreenShot
Call ClickButton("Next Page",False)

Call SetTableData("SAPLWRF_ARTICLE_SCREENSTC_AH_ASSIGNM","Main Assignment",1,"<NA>","<NA>",DT_MM41_2021_TABLECELL_MAIN_ASSIGNMENT_0,False)
Call SetTableData("SAPLWRF_ARTICLE_SCREENSTC_AH_ASSIGNM","Hierarchy Node",1,"<NA>","<NA>",DT_MM41_2021_TABLECELL_HIERARCHY_NODE_0,False)
Call SetTextbox("Brand","MARA-ZZBRAND","",DT_MM41_0001_BRAND,false)
Call SetTextbox("Sub Brand","MARA-ZZSUB_BRAND","",DT_MM41_0001_SUB_BRAND,false)
Call TakeScreenShot
Call ClickButton("LEGACY_PUSH",False)   'Extended data

Call SelectCheckbox("ZMDAM_MARAEXT-FAIR_TRADE",0,DT_MM41_4007_FAIR_TRADE,False)
Call SelectCheckbox("ZMDAM_MARAEXT-PRESENCE_OF_BBD",0,DT_MM41_4007_PRESENCE_OF_BBD,False)
Call SetTextbox("Unit weight \(gr\)","ZMDAM_MARAEXT-UNIT_WEIGHT",0,DT_MM41_4007_UNIT_WEIGHT__GR,false)
'Call SetTextbox("Unit weight \(gr\)","ZMDAM_MARAEXT-PRIVATE_BRAND",3,DT_MM41_4007_UNIT_WEIGHT__GR_OCC1,false)
Call TakeScreenShot

Call SelectCellGuiTable("SAPLZMDAM_EXT_SCRTC_LTEXT","Text","LangNu","0",False)
Call ClickCellTable("SAPLZMDAM_EXT_SCRTC_LTEXT","#1",1,"<NA>","<NA>",False)
Call SelectCellGuiTable("SAPLZMDAM_EXT_SCRTC_LTEXT","Text","LangNu","0",False)
Call TakeScreenShot
Call ClickButton("Create text",False)
Call TakeScreenShot

Call SetComboByKey("GV_DESC_LANGU_NEW",DT_MM41_0005_CREATE_TEXT_IN)
Call TakeScreenShot
Call ClickButton("Enter   \(Enter\)",True)

Call SetTextArea(DT_MM41_4007_TEXTEDIT_SHELL)
Call ClickButton("Create text",False)
Call TakeScreenShot

Call SetComboByKey("GV_DESC_LANGU_NEW",DT_MM41_0005_CREATE_TEXT_IN_OCC1)
Call SetComboByKey("GV_DESC_LANGU_NEW_VORLAGE",DT_MM41_0005_COPY_FROM)
Call TakeScreenShot
Call ClickButton("Enter   \(Enter\)",True)
Call TakeScreenShot

Call ClickButton("Back   \(F3\)",False)
Wait(2)
Call SelectTab("TABSPR1"," Purchasing",False)

Call SelectCheckbox("EINA-RELIF",0,DT_MM41_2221_REGULAR_VENDOR,False)
Call SetTextbox("Net Price","EINE-NETPR","",DT_MM41_2223_NET_PRICE,false)
Call SetTextbox("Available from","EINA-LIFAB","",ConvertDate(DT_MM41_2221_AVAILABLE_FROM),false)
Call SetTextbox("Available to","EINA-LIFBI","",ConvertDate(DT_MM41_2221_AVAILABLE_TO),false)
Call SetTextbox("Vendor artl no.","EINA-IDNLF","",DT_MM41_2221_VENDOR_ARTL_NO,false)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot
Call ClickButtonIfExist("Cancel",True)
Call ClickButtonIfExist("Cancel   \(F12\)",True)

Call ClickButton("KONDI_PUSH",False)
Call TakeScreenShot
Call VerifyTableCellContent(1,"CnTy","SAPMV13ATCTRL_D0201",DT_MM41_0201_CHECK_TEXT_OF_TABLECELL_CNTY_0)
Call VerifyTableCellContent(1,"Name","SAPMV13ATCTRL_D0201",Lcase(DT_MM41_0201_CHECK_TEXT_OF_TABLECELL_NAME_0))

Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call SetTextbox("Minimum Qty","EINE-MINBM","",DT_MM41_2222_MINIMUM_QTY,false)
Call SetTextbox("Var. Order Unit","EINA-VABME","",DT_MM41_2221_VAR_ORDER_UNIT,false)
Call TakeScreenShot

Call ClickButton("Go to additional data   \(Ctrl\+F6\)",False)    ''Additional data
Call SelectCellGuiTable("SAPLMGD2TC_KTXT","Material Description","Language","ID",False)''need to check
Call ClickButton("KT_DELETE",False)
Call ClickButton("KT_DELETE",False)
Call ClickButton("KT_DELETE",False)
Call ClickButton("KT_DELETE",False)
Call SetTableData("SAPLMGD2TC_KTXT","Material Description",1,"<NA>","<NA>",DT_MM41_8000_TABLECELL_ARTICLE_DESCRIPTION_0,False)
Call SetTableData("SAPLMGD2TC_KTXT","Material Description",2,"<NA>","<NA>",DT_MM41_8000_TABLECELL_ARTICLE_DESCRIPTION_1,False)
Call TakeScreenShot

Call SelectTab("TABSPR1","Additional GTINs",False)
Call SetTableData("SAPLMGD2TC_EAN","AV",2,"<NA>","<NA>",DT_MM41_8025_TABLECELL_AV_1,False)
Call SetTableData("SAPLMGD2TC_EAN","MV",2,"<NA>","<NA>",DT_MM41_8025_TABLECELL_MV_1,False)
Call TakeScreenShot

Call ClickButton("Go to main data   \(Ctrl\+Shift\+F3\)",False)   ''Main data
Call TakeScreenShot
Call SelectTab("TABSPR1","POS",False)
Call TakeScreenShot
Call SetTableData("SAPLMGD2TC_BON","L",1,"<NA>","<NA>",DT_MM41_2273_TABLECELL_L_0,False)
Call SetTableData("SAPLMGD2TC_BON","L",2,"<NA>","<NA>",DT_MM41_2273_TABLECELL_L_1,False)
Call SetTableData("SAPLMGD2TC_BON","AUn",1,"<NA>","<NA>",DT_MM41_2273_TABLECELL_AUN_0,False)
Call SetTableData("SAPLMGD2TC_BON","AUn",2,"<NA>","<NA>",DT_MM41_2273_TABLECELL_AUN_1,False)
Call SetTableData("SAPLMGD2TC_BON","Till rcpt texts for unit",1,"<NA>","<NA>",DT_MM41_2273_TABLECELL_TILL_RCPT_TEXTS_FOR_UNIT_0,False)
Call SetTableData("SAPLMGD2TC_BON","Till rcpt texts for unit",2,"<NA>","<NA>",DT_MM41_2273_TABLECELL_TILL_RCPT_TEXTS_FOR_UNIT_1,False)
Call TakeScreenShot

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetStatusBar("item1","DT_MM41_0100_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_MM41_0100_CHECK_TEXT_OF_STATUSBAR_OCC1)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
'
'''''--------TransactionCode-ZMDAS_WSL11----------''''
'
Call SetTcode(DT_MM41_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_MM41_0100_OKCD)

Call SetTextbox("Assortment","S_FILIA-LOW","",DT_MM41_1000_ASSORTMENT,False)
Call SetTextbox("Article","S_ARTNR-LOW","",DT_MM41_1000_ARTICLE,False)
Call SetTextbox("Customer No\. - Site","S_LOCNR-LOW","",DT_SITE_FROM,False)
Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Execute   \(F8\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectColumnGuiGrid("","","Customer No. - Site",False)
Call ClickButtonIfExist("Sort in Ascending Order   \(Ctrl\+F4\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyGridCellContent("",1,"Article","",DT_MM41_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ARTNR)
Call VerifyGridCellContent("",1,"Customer No. - Site","",DT_MM41_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOCNR)
Call VerifyGridCellContent("",1,"Assortment","",DT_MM41_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_FILIA)
Call VerifyGridCellContent("",1,"Valid From","",DT_MM41_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DATAB)
'
'''''--------TransactionCode-MM42----------''''
''Call SetTcode(DT_MM41_0500_OKCD)     
''Call PressEnter()   
''
''Call SetTextbox("Article","RMMW1-MATNR","",DT_MM41_0100_ARTICLE,False)
''Call SetTextbox("Sales Org.","RMMW1-VKORG","",DT_MM41_0100_SALES_ORG,False)
''Call SetTextbox("Distr. Channel","RMMW1-VTWEG","",DT_MM41_0100_DISTR_CHANNEL,False)
''Call SetTextbox("Purchasing Org\.","RMMW1-EKORG","","",False)
''''Call SetTextbox("Vendor","RMMW1-LIFNR","","",False)
''
''Call SetTextboxNoLabel("RMMW1-LIFNR","","",False)
''Call ClickButton("Deselect All   \(Shift\+F7\)",False)
''Call SelectRowGuiTableByRow("SAPLMGMWTAB_CONT_0100",4,False)
''Call TakeScreenShot
''Call PressEnter()
''Call TakeScreenShot
''
''Call SetTextbox("Sales unit","MVKE-VRKME","","",False)
''Call TakeScreenShot
''Call ClickButton("MVKE_PUSH",False)  ''other sales data
''Call TakeScreenShot
''Call SetTextbox("Article pricing grp","MVKE-KONDM","",DT_MM41_2152_ARTICLE_PRICING_GRP,False)
''Call TakeScreenShot
''Call ClickButton("Go to main data   \(Ctrl\+Shift\+F3\)",False)  ''main data
''Call TakeScreenShot
''Call ClickButton("Switch area of validity   \(Shift\+F1\)",False)
''Call SetTextbox("Distr. Channel","RMMW1-VTWEG","",DT_MM41_0081_DISTR_CHANNEL,True)
''Call SetTextbox("Sales Org.","RMMW1-VKORG","",DT_MM41_0081_SALES_ORG,True)
''Call SetTextbox("Sales unit ","RMMW1-VRKME","",DT_MM41_0081_SALES_UNIT,True)
''Call TakeScreenShot
''Call ClickButtonIfExist("Continue   \(Enter\)",True)
''Call TakeScreenShot
''Call SetTextbox("Final price","CALP-ENDPR","",DT_MM41_2233_FINAL_PRICE,False)
''   
''Call TakeScreenShot
''Call PressEnter() 
''Call TakeScreenShot
''Call VerifyStatusBar(DT_MM41_4030_CHECK_TEXT_OF_STATUSBAR)
''
''Call ClickButton("Save   \(Ctrl\+S\)",False)
''Wait(2)
''Call TakeScreenShot
''Call VerifyStatusBar(DT_MM41_0100_CHECK_TEXT_OF_STATUSBAR_OCC2)
'''
'''----------------------Tcode VKP5 - Setting sales price steps is changed from MM42 to VKP5 ----------------------------
Call SetTcode(DT_ZMDAS_WSL11_0500_OKCD_1) 
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

'''--------TransactionCode-VKP2----------''''

Call SetTcode(DT_MM41_0100_OKCD_OCC1)     
Call PressEnter() 

Call SetTextbox("Article","S_MATNR-LOW","",DT_MM41_1000_ARTICLE_OCC1,False)
Call SetTextbox("Sales organization","S_VKORG-LOW","",DT_MM41_1000_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution channel","S_VTWEG-LOW","",DT_MM41_1000_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Validity Period","S_DATUM-LOW","",DT_MM41_1000_VALIDITY_PERIOD,False)
Call TakeScreenShot
Call PressEnter()   

Call ClickButton("Execute   \(F8\)",False)
Wait(2)
Call TakeScreenShot

Call VerifyGridCellContent("Sales Price Conditions", 1, "Article", 0, DT_MM41_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR)
Call VerifyGridCellContent("Sales Price Conditions", 1, "Condition unit", 0, DT_MM41_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VRKME)
Call VerifyGridCellContent("Sales Price Conditions", 1, "Sales Organization", 0, DT_MM41_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VKORG)
Call VerifyGridCellContent("Sales Price Conditions", 1, "Valid from", 0, DT_MM41_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DATAB_OCC1)
Call VerifyGridCellContent("Sales Price Conditions", 1, "Condition rate", 0, DT_MM41_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KBETR)
Call VerifyGridCellContent("Sales Price Conditions", 1, "Condition Type", 0, DT_MM41_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KSCHL)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

