
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Perform an Expired Push in BW01 or BW08
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


gstrTestCaseName = "Test_Perform an Expired Push in BW01 or BW08"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\DLL\RETAIL\DT_Perform a Zero Price Push Order in BW01 or BW08_TASE.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()
Call TakeScreenShot()
''
''--------------------------------------------  MIGO ----------------------------------------------
'
'
Call SetComboByKey("GODYNPRO-ACTION",DT_MIGO_0010_GODYNPROACTION)
Call SetComboByKey("GODYNPRO-REFDOC",DT_MIGO_0010_GODYNPROREFDOC)
Call SetTextboxNoLabel("GODEFAULT_TV-BWART",0,DT_MIGO_0010_GODEFAULT_TVBWART,False)
'Call SetTextbox("Receipt w/o PO","GODEFAULT_TV-BWART",0,DT_MIGO_0010_GODEFAULT_TVBWART,False)
Call TakeScreenShot()

Call SetTextbox("Article","GODYNPRO-MAKTX",0,DT_MIGO_0390_ARTICLE,False)
Call SetTextbox("Site","GODYNPRO-NAME1",0,DT_MIGO_0390_SITE,False)
Call SetTextbox("Stor\. Loc\.","GODYNPRO-LGOBE",0,DT_MIGO_0390_STOR_LOC,False)
Call SetTextboxNoLabel("GOITEM-UMLGOBE",0,DT_MIGO_0390_GOITEMUMLGOBE,False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call SetTextbox("Qty in UnE","GODYNPRO-ERFMG",0,DT_MIGO_0390_QTY_IN_UNE,False)
Call PressEnter()

Call ClickButton("Next Item",False)
Call TakeScreenShot()

Call SetTextbox("Article","GODYNPRO-MAKTX",0,DT_MIGO_0390_ARTICLE_OCC1,False)
Call SetTextbox("Site","GODYNPRO-NAME1",0,DT_MIGO_0390_SITE_OCC1,False)
Call SetTextbox("Stor\. Loc\.","GODYNPRO-LGOBE",0,DT_MIGO_0390_STOR_LOC_OCC1,False)
Call SetTextboxNoLabel("GOITEM-UMLGOBE",0,DT_MIGO_0390_GOITEMUMLGOBE_OCC1,False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()
Call SetTextbox("Qty in UnE","GODYNPRO-ERFMG",0,DT_MIGO_0390_QTY_IN_UNE_OCC1,False)

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot()

Call GetStatusBar("item1","DT_ARTICLE_DOC_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_ARTICLE_DOC_OUTPUT",DT_ARTICLE_DOC)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR_OCC1)

Call SetTcode(DT_MIGO_0001_OKCD)
Call PressEnter()
Call TakeScreenShot()

'--------------------------------------------  WA01 ----------------------------------------------

Call SetTextbox("Allocation Table Type","AUKO-AUFAR",0,DT_MIGO_0100_ALLOCATION_TABLE_TYPE,False)
Call SetTextbox("Purchasing Organization","AUKO-EKORG",0,DT_MIGO_0100_PURCHASING_ORGANIZATION,False)
Call SetTextbox("Purchasing Group","AUKO-EKGRP",0,DT_MIGO_0100_PURCHASING_GROUP,False)

Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()

Call SetTextbox("Allocation Table","AUKO-BEZCH",0,DT_MIGO_0122_ALLOCATION_TABLE,False)
Call SetTextbox("Site Delivery Date","RM01A-WEFDT",0,ConvertDate(DT_MIGO_0122_SITE_DELIVERY_DATE),False)
Call PressEnter()
Call TakeScreenShot()

Call SelectMenuBar("Allocation table;Create with Reference;User Exit Article Selection")
Call SetTextbox("Directory","DY_PATH",0,DT_MIGO_0200_DIRECTORY,True)
Call SetTextbox("File Name","DY_FILENAME",0,DT_MIGO_0200_FILE_NAME,True)

Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot()
Call PressEnter()
Call ClickButton("Select all   \(Shift\+F8\)",False)
Call ClickButton("Fast item change   \(Ctrl\+F9\)",False)
Call TakeScreenShot()

Call SelectRadioButton("G_DATEN_ALLE_POS","Also overwrite existing values",True)
Call SetTextbox("Recipient Determination","AUPO-NREMFIN",0,DT_MIGO_0305_RECIPIENT_DETERMINATION,True)
Call SetTextbox("Allocation Strategy","AUPO-ASTRA",0,DT_MIGO_0305_ALLOCATION_STRATEGY,True)
Call SetTextbox("Item category","AUPO-APSTP",0,DT_MIGO_0305_ITEM_CATEGORY,True)
Call TakeScreenShot()
Call ClickButton("Check   \(Enter\)",True)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",True)
Call TakeScreenShot()

'Call ClickButtonIfExist("Continue   \(Enter\)",True)
'Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call VerifyTextBoxContent("Information Message","MESSTXT1",0,lcase(DT_MIGO_0010_CHECK_TEXT_OF_MESSTXT1),True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButton("Save",False)

Call GetStatusBar("item1","DT_ALLOCATION_TABEL_NUMBER_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_ALLOCATION_TABEL_NUMBER_OUTPUT",DT_ALLOCATION_TABEL_NUMBER)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_MIGO_0100_CHECK_TEXT_OF_STATUSBAR_OCC1)

Call SetTcode(DT_MIGO_0100_OKCD)
Call PressEnter()
Call TakeScreenShot()

'--------------------------------------------  WA08 ----------------------------------------------

Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Variant","V-LOW",0,DT_MIGO_0100_VARIANT,True)
Call SetTextbox("Created By","ENAME-LOW",0,"",True)

Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",True)
Call TakeScreenShot()

Call SetTextbox("Allocation Table","ABELN_RT-LOW",0,DT_MIGO_1050_ALLOCATION_TABLE,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()

Call ClickButton("Generate Documents   \(Shift\+F5\)",False)
Call TakeScreenShot()

Call GetLabelContentByRefLabel("Wrhs order",0,-32,"DT_MIGO_0120_CHECK_TEXT_OF_NO_NAME_OUTPUT",False)
Call VerifyifGuiLabelExistsByRelativeid(DT_MIGO_0120_CHECK_TEXT_OF_NO_NAME_OCC1,"wnd\[0\]/usr/lbl\[3,3\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_MIGO_0120_CHECK_TEXT_OF_NO_NAME_OCC2,"wnd\[0\]/usr/lbl\[3,4\]")

Call SetTcode(DT_MIGO_0120_OKCD)
Call PressEnter()
Call TakeScreenShot()

'--------------------------------------------  VL10G ----------------------------------------------

Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Variant","V-LOW",0,DT_MIGO_0100_VARIANT_OCC1,True)
Call SetTextbox("Created By","ENAME-LOW",0,"",True)

Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",True)
Call TakeScreenShot()

Call SelectTab("TABSTRIP_ORDER_CRITERIA","Purchase Orders",False)
Call TakeScreenShot()
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Purchasing Document","ST_EBELN-LOW",0,DT_MIGO_1030_PURCHASING_DOCUMENT,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
''Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot()
Wait 5
'Call VerifyStatusBar(lcase(DT_MIGO_1000_CHECK_TEXT_OF_STATUSBAR))

Call SetTcode(DT_MIGO_1000_OKCD)
Call PressEnter()
Call TakeScreenShot()

'--------------------------------------------  ME23N ----------------------------------------------

Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Pur\. Order","MEPO_SELECT-EBELN",0,DT_MIGO_0003_PUR_ORDER,True)
Call TakeScreenShot()
Call ClickButton("Other Document   \(Enter\)",True)
Call TakeScreenShot()

Call SelectTab("ITEM_DETAIL","Purchase Order History",False) 
Call TakeScreenShot()

Call ClickCellGuiGrid("",0,"Article Document",1,"","",False)
Call TakeScreenShot()

Call SelectTab("TAXI_TABSTRIP_OVERVIEW","Picking",False)
Call TakeScreenShot()
Wait 5
Call VerifyTableCellContent(1,"SLoc","SAPMV50ATC_LIPS_PICK",DT_MIGO_1104_CHECK_SLOC_0)
Call VerifyTableCellContent(2,"SLoc","SAPMV50ATC_LIPS_PICK",DT_MIGO_1104_CHECK_SLOC_1)

Call LogOff()
Call FinalStatus ()



