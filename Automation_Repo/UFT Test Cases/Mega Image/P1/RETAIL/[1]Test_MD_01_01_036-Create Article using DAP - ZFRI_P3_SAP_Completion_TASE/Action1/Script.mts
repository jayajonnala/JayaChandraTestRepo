

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_01_01_036-Create Article using DAP - ZFRI_P3_SAP_Completion
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
gstrTestCaseName = "Test_MD_01_01_036- - ZFRI_P3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_002_Create_Purchasing_Conditions_for_an_Existing_Article_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

'''''''----------------------Tcode MM42----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextbox("Article","RMMW1-MATNR","",DT_MM42_0100_ARTICLE,False)
Call TakeScreenShot()

Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100","Screen description","Basic Data",False)
Call TakeScreenShot()
Call PressEnter()
wait(2)

Call VerifyTextBoxContent("Supply source", "MARA-BWSCL", 0, DT_MM42_2003_CHECK_TEXT_OF_SUPPLY_SOURCE, False)
Call SetTextbox("Valuation Class","MAW1-WBKLA","",DT_MM42_2002_VALUATION_CLASS,False)
Call SetTextbox("Packag\.mat\.type","MARA-VHART","",DT_MM42_2164_PACKAGMATTYPE,False)
Call SetTextboxNoLabel("MARA-MSTAE","","",False)
Call SetTextbox("Valid from","MARA-MSTDE","","",False)
Call SetTextbox("GenItemCatGroup","MARA-MTPOS_MARA","",DT_MM42_2003_GENITEMCATGROUP,False)
Call SetTextbox("Brand","MARA-ZZBRAND","",DT_MM42_0001_BRAND,False)
Call PressEnter()
wait(2)
Call TakeScreenShot()
Call ClickButtonIfExist("Cancel",True)
Call TakeScreenShot()

Call SelectTab("TABSPR1","Sales", False)
Call TakeScreenShot()
Call ClickButton("Switch area of validity   \(Shift\+F1\)",False)
Call TakeScreenShot()

Call SetTextbox("Sales Unit","RMMW1-VRKME","",DT_MM42_0081_SALES_UNIT,True)
Call SetTextbox("Sales Org\.","RMMW1-VKORG","",DT_MM42_0081_SALES_ORG,True)
Call SetTextbox("Distr\. Channel","RMMW1-VTWEG","",DT_MM42_0081_DISTR_CHANNEL,True)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot()

Call SetTextbox("Item cat\.group","MVKE-MTPOS","",DT_MM42_2231_ITEM_CATGROUP,False)
Call TakeScreenShot()
Call ClickButton("Other Sales Data",False)
Call TakeScreenShot()

Call SetTextbox("Article pricing grp","MVKE-KONDM","",DT_MM42_2152_ARTICLE_PRICING_GRP,False)
Call SetTextbox("Acct assignment grp","MVKE-KTGRM","",DT_MM42_2152_ACCT_ASSIGNMENT_GRP,False)
Call SetTextbox("Item category group","MVKE-MTPOS","",DT_MM42_2152_ITEM_CATEGORY_GROUP,False)
Call SetTextbox("Gen\. item cat\. grp","MARA-MTPOS_MARA","",DT_MM42_2152_GEN_ITEM_CAT_GRP,False)

Call PressEnter()
wait(2)
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False)

Call SetTextbox("Item Set","MVKE-MVGR4","",DT_MM42_2154_ITEM_SET,False)
Call ClickBUtton("Save   \(Ctrl\+S\)",False)
Call verifyStatusbar(Lcase(DT_MM42_0100_CHECK_TEXT_OF_STATUSBAR))


''''----------------------Tcode  VKP5 ----------------------------
'''Enter the transaction code
Call SetTcode(DT_MM42_0100_OKCD) 
Call PressEnter()     ' 
Call TakeScreenShot

Call ClickButton("Get variant\.\.\.   \(Shift\+F5\)",False)
Call TakeScreenShot
Call SetTextbox("Created By","ENAME-LOW","","",False)
Call SetTextbox("Variant","V-LOW","",DT_MM42_0100_VARIANT,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",True)
Wait 2

Call SetTextbox("Article","S_MATNR-LOW","",DT_MM42_1000_ARTICLE,False)
Call SetTextbox("Sales organization","S_VKORG-LOW","",DT_MM42_1000_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution channel","S_VTWEG-LOW","",DT_MM42_1000_DISTRIBUTION_CHANNEL,False)
Call PressEnter()     ' 
Call TakeScreenShot

Call ClickButton("Execute   \(F8\)",FalsE)
Call takeScreenShot()

Call GetGridContentByTitle("", 1, "Final price", 1, "DT_MM42_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ENDPR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_MM42_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ENDPR_OUTPUT",DT_MM42_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ENDPR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetGridData("", 1, "ENDPR", DT_MM42_0100_GRIDCELL_0_FINAL_PRICE, False)
Call takeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call takeScreenShot()
Call GetStatusBar("item1","DT_PRICING_DOC_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_PRICING_DOC_OUTPUT",DT_PRICING_DOC)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

CAll GetTextStatusBar("DT_MM42_1000_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar(Lcase(DT_MM42_1000_CHECK_TEXT_OF_STATUSBAR_OUTPUT))
Call WriteRunTimeDataToExcelGlobalSheet("DT_MM42_1000_CHECK_TEXT_OF_STATUSBAR_OUTPUT",DT_MM42_1000_CHECK_TEXT_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

''''----------------------Tcode  WSM3 ----------------------------
'''Enter the transaction code
Call SetTcode(DT_MM42_1000_OKCD) 
Call PressEnter()     ' 
Call TakeScreenShot

Call SetTextbox("Article","MATNR-LOW","",DT_MM42_1000_ARTICLE_OCC1,False)
Call SetTextbox("Listing Procedures","LSTFL","",DT_MM42_1000_LISTING_PROCEDURES,False)
Call SetTextbox("Assortment","ASORT-LOW","",DT_MM42_3010_TABLECELL_SINGLE_VALUE_0_OCC1,False)
Call TakeScreenShot
Call ClickButton("%_ASORT_%_APP_%-VALU_PUSH",False)
Call SetTableDataNoRef("SAPLALDBSINGLE", "Single value", 1, DT_MM42_3010_TABLECELL_SINGLE_VALUE_0_OCC1, False)
Call SetTableDataNoRef("SAPLALDBSINGLE", "Single value", 2, DT_MM42_3010_TABLECELL_SINGLE_VALUE_1_OCC1, False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Copy   \(F8\)",False)

Call SelectCheckbox("DATBE",0,"ON", False)
Call SelectCheckbox("LSTFLMAT", 0, DT_MM42_1000_NOTE_ARTICLE_LISTING_PROC, False)
Call SelectCheckbox("NEWLIST", 0, DT_MM42_1000_NEW_LISTING, False)
Call SelectCheckbox("LIEFWERK", 0, DT_MM42_1000_ALSO_LIST_SUPPLYING_SITE, False)
Call SelectCheckbox("PROT_OUT", 0, DT_MM42_1000_ISSUE_LOG, False)
Call SelectCheckbox("NO_DOCU", 0, DT_MM42_1000_DEACTIVATE_CHANGE_DOCUMENTS, False)
Call SelectCheckbox("CHECK_LO", 0, DT_MM42_1000_LIST_LOCAL_ASSORTMENTS_ALSO, False)

Call ClickBUtton("Execute   \(F8\)",False)

Call VerifyifGuiLabelExists(DT_MM42_0120_CHECK_TEXT_OF_RW04)
Call VerifyifGuiLabelExists(DT_MM42_0120_CHECK_TEXT_OF_RW21)
Call VerifyifGuiLabelExistsByRelativeid(DT_MM42_0120_CHECK_TEXT_OF_NO_NAME, "wnd\[0\]/usr/lbl\[1,9\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_MM42_0120_CHECK_TEXT_OF_Z03020301, "wnd\[0\]/usr/lbl\[12,8\]")

Call VerifyifGuiLabelExistsByRelativeid(DT_MM42_0120_CHECK_TEXT_OF_NO_NAME_OCC1, "wnd\[0\]/usr/lbl\[1,12\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_MM42_0120_CHECK_TEXT_OF_Z03020301_OCC1, "wnd\[0\]/usr/lbl\[12,11\]")


'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
