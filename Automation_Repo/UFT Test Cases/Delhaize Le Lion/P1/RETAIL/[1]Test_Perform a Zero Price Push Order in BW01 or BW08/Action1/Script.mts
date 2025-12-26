
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Perform a Zero Price Push in BW01&BW08
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

gstrTestCaseName = "Test_Perform a Zero Price Push in BW01&BW08"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\DLL\RETAIL\DT_Perform a Zero Price Push Order in BW01 or BW08_TASE.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'all LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''''
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call TakeScreenShot()

'--------------------------------------------  WA01 ----------------------------------------------

Call SetTextbox("Allocation Table Type","AUKO-AUFAR","",DT_WA01_0100_ALLOCATION_TABLE_TYPE,False)
Call SetTextbox("Purchasing Organization","AUKO-EKORG","",DT_WA01_0100_PURCHASING_ORGANIZATION,False)
Call SetTextbox("Purchasing Group","AUKO-EKGRP","",DT_WA01_0100_PURCHASING_GROUP,False)
Call TakeScreenShot()

' Added lines 51, 58, 61, 64, 68, 71-73, 81, 85, 88-91, 95-101, 137-145, 155-171, 175-184
' By ARUD on 16th Mar 2022

Call PressEnter() 

Call TakeScreenShot()
Call SetTextbox("Site Delivery Date","RM01A-WEFDT","",ConvertDate(DT_WA01_0122_SITE_DELIVERY_DATE),False)
Call SetTextbox("Allocation Table","AUKO-BEZCH","",DT_WA01_0122_ALLOCATION_TABLE,False)
Call TakeScreenShot()
Call PressEnter() 

Call TakeScreenShot()
call SelectMenuBar("Allocation table;Create with Reference;User Exit Article Selection")
Call SetTextbox("Directory","DY_PATH","",DT_WA01_0200_DIRECTORY,False)
Call SetTextbox("File Name","DY_FILENAME","",DT_WA01_0200_FILE_NAME,False)
Call TakeScreenShot()

Call ClickButtonIfExist("Continue   \(Enter\)",False)
Call TakeScreenShot()
Call PressEnter()     
Call VerifyTableCellContent(1,"Item category","SAPML01ATC_0122",DT_WA01_0122_CHECK_TEXT_OF_TABLECELL_ITEM_CATEGORY_0)
Call VerifyTableCellContent(2,"Item category","SAPML01ATC_0122",DT_WA01_0122_CHECK_TEXT_OF_TABLECELL_ITEM_CATEGORY_1)

Call ClickButton("Select all   \(Shift\+F8\)",False)
Call ClickButton("Fast item change   \(Ctrl\+F9\)",False)

Call SelectRadioButton("G_DATEN_ALLE_POS","Also overwrite existing values",True)
Call SetTextbox("Recipient Determination","AUPO-NREMFIN","",DT_WA01_0305_RECIPIENT_DETERMINATION,False)
Call SetTextbox("Allocation Strategy","AUPO-ASTRA	","",DT_WA01_0305_ALLOCATION_STRATEGY,False)
Call SetTextbox("Item category","AUPO-APSTP","",DT_WA01_0305_ITEM_CATEGORY,False)
Call TakeScreenShot()

Call ClickButtonIfExist("Continue   \(Enter\)",False)
Call ClickButtonIfExist("Check   \(Enter\)",false)
Call TakeScreenShot()
Call ClickButtonIfExist("Execute   \(F8\)",false)

'Call ClickButtonIfExist("Continue   \(Enter\)",True)
'Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call VerifyTextBoxContent("Information Message","MESSTXT1",0,lcase(DT_WA01_0010_CHECK_TEXT_OF_MESSTXT1),True)
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButton("Save",True)

Call GetStatusBar("item1","DT_ALLOCATION_TABLE_NUMBER_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_ALLOCATION_TABLE_NUMBER_OUTPUT",DT_ALLOCATION_TABLE_NUMBER)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(lcase(DT_WA01_0100_CHECK_TEXT_OF_STATUSBAR_OCC1))


Call SetTcode(DT_WA01_0100_OKCD) 
Call PressEnter()     ' 
Call TakeScreenShot()

'--------------------------------------------  WA08 ----------------------------------------------

Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Variant","V-LOW","",DT_WA01_0100_VARIANT,True)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call TakeScreenShot()
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call SetTextbox("Allocation Table","ABELN_RT-LOW","",DT_WA01_1050_ALLOCATION_TABLE,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()

Call ClickButton("Generate Documents   \(Shift\+F5\)",False)
Call TakeScreenShot()
Call SetHorizontalScrollBar(50,False)
Call GetLabelContentByRefLabel("Sales ord.",0,-32,"DT_WA01_0120_CHECK_TEXT_OF_NO_NAME_OUTPUT",False)
Call VerifyifGuiLabelExistsByRelativeid(DT_WA01_0120_CHECK_ICONNAME_OF_NO_NAME,"wnd\[0\]/usr/lbl\[3,3\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_WA01_0120_CHECK_ICONNAME_OF_NO_NAME_OCC1,"wnd\[0\]/usr/lbl\[3,4\]")

Call SetTcode(DT_WA01_0120_OKCD)
Call PressEnter()
Call TakeScreenShot()

'--------------------------------------------  VL10G ----------------------------------------------

Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Variant","V-LOW","",DT_WA01_0100_VARIANT_OCC1,True)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",True)
Call TakeScreenShot()
Call SelectTab ("TABSTRIP_ORDER_CRITERIA","Sales Orders",False)
Call TakeScreenShot()
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextboxNoLabel("ST_VBELN-LOW","",DT_WA01_1020_SALES_DOCUMENT,False)
Call PressEnter()
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()
''Call VerifyStatusBar(lcase(DT_WA01_1000_CHECK_TEXT_OF_STATUSBAR))

Call SetTcode(DT_WA01_0120_OKCD_OCC1)
Call PressEnter()
Call TakeScreenShot()
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'--------------------------------------------  VA03 ----------------------------------------------

Call SetTextbox("Order","VBAK-VBELN","",DT_WA01_0102_ORDER,False)
Call PressEnter()
Call TakeScreenShot()
Call VerifyTableCellContent(1,"ItCa","SAPMV45ATCTRL_U_ERF_AUFTRAG",DT_WA01_4900_CHECK_TEXT_OF_TABLECELL_ITCA_0)
call VerifyTableCellContent(2,"ItCa","SAPMV45ATCTRL_U_ERF_AUFTRAG",DT_WA01_4900_CHECK_TEXT_OF_TABLECELL_ITCA_1)
Call ClickButton("Display header details",False)
Call VerifyTextBoxContent("Order Type","VBAK-AUART","",DT_WA01_4301_CHECK_TEXT_OF_ORDER_TYPE,False)
Call TakeScreenShot()
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Display document flow   \(F5\)",False)
Call TakeScreenShot()

Call SelectNodeGuiTree(0,"#1;#1;#1")
Call ClickButton("Display document   \(F8\)",False)
Call TakeScreenShot()

Call GetTextboxValue("LIKP-VBELN",0,DT_WA01_1502_CHECK_TEXT_OF_OUTBOUND_DELIV_Output,False)
Call VerifyTableCellContent(1,"Article","SAPMV50ATC_LIPS_OVER",DT_WA01_1102_CHECK_TEXT_OF_TABLECELL_ARTICLE_0)
Call VerifyTableCellContent(2,"Article","SAPMV50ATC_LIPS_OVER",DT_WA01_1102_CHECK_TEXT_OF_TABLECELL_ARTICLE_1)

Call SetTcode(DT_WA01_1000_OKCD)
Call PressEnter()
Call TakeScreenShot()
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'--------------------------------------------  VL02N ----------------------------------------------

Call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_WA01_4004_OUTBOUND_DELIVERY,False)
Call SelectMenuBar("Subsequent Functions;Output from Deliveries")
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)

' selected=True, Target Process button, verfy statusbar must be added
Call TakeScreenShot()

Call LogOff()
Call FinalStatus ()



