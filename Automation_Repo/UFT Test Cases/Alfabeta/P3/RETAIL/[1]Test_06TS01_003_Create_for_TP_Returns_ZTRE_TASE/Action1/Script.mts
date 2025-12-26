

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_06TS01_003_Create_for_TP_Returns_ZTRE
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
	GetRowNo= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)


gstrTestCaseName = "Test_06TS01_003_Returns_ZTRE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''
'''--------------------------------ME21N-----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetComboByKey("MEPO_TOPLINE-BSART",DT_ME21N_1105_MEPO_TOPLINEBSART)
'Call SetTextbox("Vendor","MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,False)   
Call SetTextboxNoLabel("MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,False)
Call SetTextbox("Purch. Org.","MEPO1222-EKORG","",DT_ME21N_1221_PURCH_ORG,False) 
Call SetTextbox("Purch. Group","MEPO1222-EKGRP","",DT_ME21N_1221_PURCH_GROUP,False)     
Call SetTextbox("Company Code","MEPO1222-BUKRS","",DT_ME21N_1221_COMPANY_CODE,False)  
Call TakeScreenSHot()
Call PressEnter()
Call TakeScreenSHot()
Call SetTableData("SAPLMEGUITC_1211","Returns Item","1","","",DT_ME21N_1211_TABLECELL_RETURNS_ITEM_0,False)
Call SetTableData("SAPLMEGUITC_1211","Returns Item","2","","",DT_ME21N_1211_TABLECELL_RETURNS_ITEM_1,False)

Call SetTableData("SAPLMEGUITC_1211","Article","1","","",DT_ME21N_1211_TABLECELL_ARTICLE_0,False) 
Call SetTableData("SAPLMEGUITC_1211","Article","2","","",DT_ME21N_1211_TABLECELL_ARTICLE_1,False)

Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_0,False)
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","2","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_1,False)

Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_ME21N_1211_TABLECELL_SITE_0,False) 
Call SetTableData("SAPLMEGUITC_1211","Plnt","2","","",DT_ME21N_1211_TABLECELL_SITE_1,False)
Call TakeScreenSHot()
Call PressEnter()
Call TakeScreenSHot()
Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call TakeScreenSHot()
Call SelectTab("HEADER_DETAIL","Partners",False)
Call SetTableData("SAPLEKPATC_0111","Funct","1","","",DT_ME21N_0111_TABLECELL_FUNCT_0,False) 
Call SetTableData("SAPLEKPATC_0111","Number","1","","",DT_ME21N_0111_TABLECELL_NUMBER_0,False)
Call TakeScreenSHot()
Call PressEnter()
Call PressEnter()
Call TakeScreenSHot()
Call ClickButton("Save   \(Ctrl\+S\)",False)
wait 2
Call GetStatusBar("item2","DT_ME21N_0014_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("AB Triangular Sales created under the number " & DT_ME21N_0014_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT )
Call WriteRunTimeDataToExcelGlobalSheet ("DT_ME21N_0014_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_ME21N_0014_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
'''--------------------------------MIGO-----------------------------

Call SetTcode(DT_ME21N_0014_OKCD)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)

call SetComboByKey("GODYNPRO-ACTION",DT_ME21N_0010_GODYNPROACTION)
call SetComboByKey("GODYNPRO-REFDOC",DT_ME21N_0010_GODYNPROREFDOC)
Call SetTextboxNoLabel("GODEFAULT_TV-BWART",0,DT_ME21N_0010_GODEFAULT_TVBWART,False)
Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER",0,DT_ME21N_2000_GODYNPROPO_NUMBER,False)
Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_ME21N_0110_DELIVERY_NOTE,False)
Call TakeScreenSHot()
Call ClickButtonIfExist("MIGO_OK_GO",False)
Call TakeScreenSHot()
Call VerifyTableCellContent(1,"Article","SAPLMIGOTV_GOITEM",DT_ME21N_0200_CHECK_TEXT_OF_TABLECELL_ARTICLE_0)
Call VerifyTableCellContent(2,"Article","SAPLMIGOTV_GOITEM",DT_ME21N_0200_CHECK_TEXT_OF_TABLECELL_ARTICLE_1)
Call VerifyTableCellContent(1,"Movement type","SAPLMIGOTV_GOITEM",DT_ME21N_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_0)
Call TakeScreenSHot()
CAll SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,"ON",False)
Call SetTableData("SAPLMIGOTV_GOITEM","OK","1","","","ON",False)
Call SetTableData("SAPLMIGOTV_GOITEM","OK","2","","","ON",False)
Call TakeScreenSHot()
Call TakeScreenShot()
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenSHot()
Call ClickButtonIfExist("Save",True)  
Call TakeScreenSHot()
Call GetStatusBar("item1","DT_ME21N_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Article document "& DT_ME21N_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT &" posted" )
Call WriteRunTimeDataToExcelGlobalSheet ("DT_ME21N_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_ME21N_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

''--------------------------------ZMDSO_TRIANG_SO--------------------------------
Call SetTcode(DT_ME21N_0001_OKCD)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call TakeScreenSHot()
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call ClickButtonIfExist("Execute   \(F8\)",True) 
Call SelectRowGuiGrid("Variant Catalog for Program ZMDSO_TRIANGULAR_SO_FROM_GR.*","","Variant name",DT_ME21N_0600_GRIDCELL_0_VARIANT_NAME,True)
Call TakeScreenSHot()
Call ClickButtonIfExist("Choose   \(F2\)",True) 
Call TakeScreenSHot()
call SetComboByKey("P_ACTION",DT_ME21N_1000_ACTION)
Call SetTextbox("Article Document","P_MBLNR","","",False)
Call SetTextbox("Article Doc\. Year","P_MJAHR","","",False)
Call SetTextbox("Posting Date","S_BUDAT-LOW","",ConvertDate(DT_ME21N_1000_POSTING_DATE_FROM),False)
Call SetTextbox("to","S_BUDAT-HIGH","",ConvertDate(DT_ME21N_1000_POSTING_DATE_TO),False)

Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenSHot()
Call FindRowNumber("","Purchasing Document",DT_ME21N_2000_GODYNPROPO_NUMBER,"DT_ROW_NO_PO_OUTPUT")

Call VerifyGridCellContentbyName("shell",DT_ROW_NO_PO_OUTPUT,"Article Document","",DT_ME21N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MBLNR)
'Call VerifyGridCellContentbyName("shell",DT_ROW_NO_PO_OUTPUT,"Vendor","",DT_ME21N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LIFNR)
Call VerifyGridCellContentbyName("shell",DT_ROW_NO_PO_OUTPUT,"Supplier","",DT_ME21N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LIFNR)
Call TakeScreenSHot()
Call VerifyGridCellContentbyName("shell",DT_ROW_NO_PO_OUTPUT,"Sales Document Type","",DT_ME21N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AUART)
Call VerifyGridCellContentbyName("shell",DT_ROW_NO_PO_OUTPUT,"Reference","",DT_ME21N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_XBLNR)
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False)
Call SelectCheckbox("P_TEST",0,DT_ME21N_1000_TEST_MODE,False)
Call SelectCheckbox("P_DISPSO",0,DT_ME21N_1000_DISPLAY_EXISTING_SO,False)
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Call GetGridContentByRefColumn("",0,"Purchasing Document",DT_ME21N_2000_GODYNPROPO_NUMBER,"Sales Document","DT_ME21N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VBELN_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_ME21N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VBELN_OUTPUT",DT_ME21N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VBELN)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call TakeScreenShot()

'''--------------------------------va03-----------------------------
Call SetTcode(DT_ME21N_0500_OKCD) 
Call PressEnter()     
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)

Call SetTextbox("Order","VBAK-VBELN","",DT_ME21N_0102_ORDER,False)
Call TakeScreenShot()
Call ClickButton("Search",False) 
Call TakeScreenShot()

Call VerifyTextBoxContent("Purch\. Order No\.","VBKD-BSTKD","",LCASE(DT_ME21N_4021_CHECK_TEXT_OF_PO_NUMBER),False)
Call VerifyTextBoxContent("Sold-To Party","KUAGV-KUNNR","",LCASE(DT_ME21N_4701_CHECK_TEXT_OF_SOLDTO_PARTY),False)
Call VerifyTableCellContent(1,"Article","SAPMV45ATCTRL_U_ERF_AUFTRAG",DT_ME21N_4900_CHECK_TEXT_OF_TABLECELL_ARTICLE_0)
Call VerifyTableCellContent(2,"Article","SAPMV45ATCTRL_U_ERF_AUFTRAG",DT_ME21N_4900_CHECK_TEXT_OF_TABLECELL_ARTICLE_1)

Call SelectMenuBar("Goto;Header;Sales")
Call VerifyTextBoxContent("Order Type","VBAK-AUART","",UCASE(DT_ME21N_4301_CHECK_TEXT_OF_ORDER_TYPE),False)
Call SelectTab("TAXI_TABSTRIP_HEAD","Order Data",False)
Call VerifyTextBoxContent("PO Number","VBKD-BSTKD","",UCASE(DT_ME21N_4351_CHECK_TEXT_OF_PO_NUMBER),False)
Call VerifyTextBoxContent("Collective no\.","VBAK-SUBMI","",UCASE(DT_ME21N_4351_CHECK_TEXT_OF_COLLECTIVE_NO),False)
Call ClickButton("Display document flow   \(F5\)",False)
Call TakeScreenSHot()
Call ActivateItemGuiTree(0,"#1;#1","#1")
Call TakeScreenSHot()
Call GetGridContentByTitle("AB_TriangularReturns .*", 0, "Doc.no.", 1, "DT_ME21N_0100_CHECK_TEXT_OF_TREE_SAPTABLETREECONTROL1_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_ME21N_0100_CHECK_TEXT_OF_TREE_SAPTABLETREECONTROL1_OUTPUT",DT_ME21N_0100_CHECK_TEXT_OF_TREE_SAPTABLETREECONTROL1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

''--------------------------------VL06O-----------------------------
Call SetTcode(DT_ME21N_0100_OKCD) 
Call PressEnter()     
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)
Call TakeScreenSHot()
Call ClickButton("For Goods Issue",False)
Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call TakeScreenSHot()
Call ClickButtonIfExist("Execute   \(F8\)",True) 
Call SelectRowGuiGrid("Variant Catalog for Program WS_MONITOR_OUTB_DEL_GDSI.*","","Variant name",DT_ME21N_0600_GRIDCELL_3_VARIANT_NAME,True)
Call TakeScreenSHot()
Call ClickButtonIfExist("Choose   \(F2\)",True) 
Call TakeScreenSHot()
Call ClickButtonIfExist("All selections   \(Shift\+F7\)",False)
Call SetTextbox("Delivery","IT_VBELN-LOW","",DT_ME21N_1000_DELIVERY,False)
Call ClickButton("Execute   \(F8\)",False) 
Call TakeScreenSHot

Call SelectCheckboxNoLabel(1, DT_ME21N_0120_NO_NAME, False)
Call TakeScreenSHot()
Call ClickButton("Post Goods Issue   \(F8\)",False)
Call TakeScreenSHot()
Call SetTextbox("Act\. Gds Mvmnt Date","LIKP-WADAT_IST","",ConvertDate(DT_ME21N_1100_ACT_GDS_MVMNT_DATE),True)
Call ClickBUtton("Continue   \(Enter\)",True)
Call TakeScreenSHot()
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC6)
Call VerifyStatusBar(Lcase(DT_ME21N_0120_CHECK_TEXT_OF_STATUSBAR))

'''--------------------------------VF01-----------------------------
Call SetTcode(DT_ME21N_0120_OKCD) 
Call PressEnter()     
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC5)
Call TakeScreenSHot()
Call SetTableData("SAPMV60ATCTRL_ERF_FAKT","Document","1","","",DT_ME21N_0102_TABLECELL_DOCUMENT_0,False)
Call TakeScreenSHot()
Call PressEnter()

CAll VerifyTableCellContent(1, "Article", "SAPMV60ATCTRL_UEB_FAKT", DT_ME21N_0104_CHECK_TEXT_OF_TABLECELL_ARTICLE_0)
CAll VerifyTableCellContent(2, "Article", "SAPMV60ATCTRL_UEB_FAKT", DT_ME21N_0104_CHECK_TEXT_OF_TABLECELL_ARTICLE_1)
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC7)
Call GetStatusBar("item1","DT_ME21N_0102_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_ME21N_0102_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_ME21N_0102_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyStatusBar(Lcase(DT_ME21N_0102_CHECK_TEXT_OF_STATUSBAR))

'''--------------------------------VA03-----------------------------
Call SetTcode(DT_ME21N_0102_OKCD) 
Call PressEnter()     
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC6)

Call SetTextbox("Order","VBAK-VBELN","",DT_ME21N_0102_ORDER_OCC1,False)
Call TakeScreenShot()
Call ClickButton("Search",False) 
Call TakeScreenShot()
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC8)
Call ClickButton("Display document flow   \(F5\)",False)

Call ActivateItemGuiTree(0,"#1;#1;#1","#1")
Call ClickBUtton("Display document   \(F8\)",False)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC7)
Call ClickBUtton("Accounting Documents   \(F7\)",False)
Call TakeScreenSHot()
CAll DoubleClickGuiGridCell("Documents in Accounting", 0, 1, "Doc. Number",True)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC8)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_ME21N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_ME21N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyTextBoxContent("Reference","BKPF-XBLNR", 0, DT_ME21N_0750_CHECK_TEXT_OF_REFERENCE, False)


Call ClickButton("Back   \(F3\)",False)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC9)
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Back   \(F3\)",False)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC10)

Call ActivateItemGuiTree(0,"#1;#1;#2;#1","#1")
Call ClickBUtton("Display document   \(F8\)",False)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC11)

Call VerifyGridCellContent("", 1, "KTONR", 0, DT_ME21N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR_OCC1)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_ME21N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR_OCC1)
Call VerifyGridCellContent("", 3, "KTONR", 0, DT_ME21N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)

Call LogOff()
Call FinalStatus ()

