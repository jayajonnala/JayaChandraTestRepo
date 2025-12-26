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

 '.......................Mandatory Initial Call only in First Component in a Test Scenario
'reload DS to update dates and calculations
''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Intracompany Store returns to DC - SW31_p1_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 19th June
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Intracompany Store returns to DC - SW31_p1_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\DS\Retail\DT_Intracompany Store returns to DC - SW31_p1_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'Increment the parameter
'Call WriteRunTimeDataToExcelGlobalSheet ("DT_ME21N_1211_CHECK_TEXT_OF_TABLECELL_REQMT_NO_0",(Cint(DT_ME21N_1211_CHECK_TEXT_OF_TABLECELL_REQMT_NO_0)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'''----------------------Tcode ME21N----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

''Enter the Details
Call SetComboByKey("MEPO_TOPLINE-BSART",DT_ME21N_1105_MEPO_TOPLINEBSART)
wait(1)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call ClickButtonIfExist("Switch Off Document Overview   \(F9\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Expand Items Ctrl\+F3",False)
wait(1)
Call TakeScreenShot()
'Call SetTableData("SAPLMEGUITC_1211","Returns Item","1","","",DT_ME21N_1211_TABLECELL_RETURNS_ITEM_0,False)
Call SetTableData("SAPLMEGUITC_1211","Article","1","","",DT_ME21N_1211_TABLECELL_ARTICLE_0,False)
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_0,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_ME21N_1211_TABLECELL_SITE_0,False)
Call SetTableData("SAPLMEGUITC_1211","Reqmt No.","1","","",DT_ME21N_1211_TABLECELL_REQMT_NO_0,False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Supplying Site","MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_SUPPLYING_SITE,False)     
'Call SetTextbox("Doc\. date","MEPO_TOPLINE-BEDAT","",Replace((DT_ME21N_1105_DOC_DATE),"/","."),False) 
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
wait(1)
Call TakeScreenShot()
Call SetTextbox("Purch. Org.","MEPO1222-EKORG","",DT_ME21N_1221_PURCH_ORG,False) 
Call SetTextbox("Purch. Group","MEPO1222-EKGRP","",DT_ME21N_1221_PURCH_GROUP,False)     
Call SetTextbox("Company Code","MEPO1222-BUKRS","",DT_ME21N_1221_COMPANY_CODE,False) 
Call TakeScreenShot()
Call PressEnter()
wait(1)
Call TakeScreenShot()
'Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)
Call ClickButtonIfExist("Expand Items Ctrl\+F3",False)
wait(1)
Call TakeScreenShot()
Call GetTableCellData("SAPLMEGUITC_1211","Reqmt No.",1,"","","DT_ME21N_1211_CHECK_TEXT_OF_TABLECELL_REQMT_NO_0_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'Click on Save Buton
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Wait(1)
Call TakeScreenShot()
Call ClickButtonIfExist("Continue",True)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Save",True)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()
'save it to data sheet
'Call GetStatusBar("item1","DT_ME21N_0014_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetStatusBar("item2","DT_ME21N_0014_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
'reload data sheet
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'verify statusbar with datasheet feed
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC2)
'Call VerifyStatusBar(DT_ME21N_0014_CHECK_TEXT_OF_STATUSBAR)
'
'----------------------------Additional Code To process in vl10b for outbound delivery doc creation_need for P2 script--------------
''''Call SetTcode("VL10B") 
''''Call PressEnter()
''''Call SelectTab("TABSTRIP_ORDER_CRITERIA","Purchase Orders",False)
''''Call SetTextbox("Shipping Point/Receiving Pt","ST_VSTEL-LOW","","",False)
''''Call SetTextbox("Deliv\. Creation Date","ST_LEDAT-LOW","","",False)
''''Call SetTextbox("to","ST_LEDAT-HIGH","","",False)
''''Call SetTextbox("Purchasing Document","ST_EBELN-LOW","",DT_ME21N_0003_PUR_ORDER,False)
''''Call ClickButton("Execute   \(F8\)",False)
''''Wait(1)
''''Call SelectRowGuiGridbyRowNo("","",1,False)
''''Call ClickButton("Create Delivery in Background   \(Shift\+F7\)",False)
''''Wait(2)
'Call VerifyStatusBarMessageType("S")
'______________________________________________________________________________________________________________________
'
'''----------------------Tcode ME23N----------------------------
'Enter the Tcode
Call SetTcode(DT_ME21N_0014_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_ME21N_0014_OKCD)
'Capture the screenshot
Call TakeScreenShot()
'Enter the Details
Call ClickButtonIfExist("Other Purchase Order   \(Shift\+F5\)",False)
'Call SetTextboxPopupIfExist("MEPO_SELECT-EBELN","Pur\. Order",DT_ME23N_0003_PUR_ORDER)
Call SetTextbox("Pur\. Order","MEPO_SELECT-EBELN","",DT_ME21N_0003_PUR_ORDER,True)
Call TakeScreenShot()
Call ClickButton("Other Document   \(Enter\)",True)

'close document overview if exist
Call ClickButtonIfExist("Switch Off Document Overview   \(F9\)",False)

'Close help page if exist
If SAPGuiSession("transaction:="&DT_ME21N_0014_OKCD).SAPGuiWindow("transaction:="&DT_ME21N_0014_OKCD).Page("title:=").Exist = True Then	
	Call ClickButton("Help   \(Shift\+F1\)",False)
	Wait(1)	
End If

'navigate to item details section
Call ClickButtonIfExist("Expand Item Details Ctrl\+F4",False)
'Texts tab
Call SelectTab("ITEM_DETAIL","Texts",False)
Call TakeScreenShot()
'Purchase order history tab
Call SelectTab("ITEM_DETAIL","Purchase Order History",False)
Call TakeScreenShot()
'Call ClickButtonToolBar("&FIND","")
Call ClickButtonToolBar("&FIND",0)
'Call SetTextboxPopupIfExist("GS_SEARCH-VALUE","Search Term:",DT_ME23N_0841_SEARCH_TERM) 
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_ME21N_0841_SEARCH_TERM,True)
Call ClickButtonIfExist("OK   \(Enter\)",True)
Wait(1)
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonToolBar("&DETAIL",0)
Wait(1)
Call TakeScreenShot()
'get article document nummber
Call GetCellDataGuiGridPopupByRefTwoColumns("","","Cell Content","Group Description","Article Document","Group Description","Article Document","DT_ME21N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_VALUE_OUTPUT")
Call ClickButtonIfExist("Close window   \(Enter\)",True)
'reload data sheet
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'''----------------------Tcode VL02N----------------------------
'Enter the Tcode
Call SetTcode(DT_ME21N_0014_OKCD_OCC1) 
Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

'Enter data
Call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_ME21N_4004_OUTBOUND_DELIVERY,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Post Goods Issue   \(Shift\+F8\)",False)
'
'''----------------------Tcode ME23N----------------------------
'Enter the Tcode
Call SetTcode(DT_ME21N_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_ME21N_0100_OKCD)
'Capture the screenshot
Call TakeScreenShot()
'Enter the Details
Call ClickButtonIfExist("Other Purchase Order   \(Shift\+F5\)",False)
'Call SetTextboxPopupIfExist("MEPO_SELECT-EBELN","Pur\. Order",DT_ME23N_0003_PUR_ORDER)
Call SetTextbox("Pur\. Order","MEPO_SELECT-EBELN","",DT_ME21N_0003_PUR_ORDER_OCC1,True)
Call TakeScreenShot()
Call ClickButton("Other Document   \(Enter\)",True)

'close document overview if exist
Call ClickButtonIfExist("Switch Off Document Overview   \(F9\)",False)

'Close help page if exist
If SAPGuiSession("transaction:="&DT_ME21N_0100_OKCD).SAPGuiWindow("transaction:="&DT_ME21N_0100_OKCD).Page("title:=").Exist = True Then	
	Call ClickButton("Help   \(Shift\+F1\)",False)
	Wait(1)	
End If

Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
wait(1)

Call SelectTab("HEADER_DETAIL","Status",False)
Call TakeScreenShot()

Call SelectTab("HEADER_DETAIL","Returns",False)
Call TakeScreenShot()

Call ClickButton("Open Returns Overview",False)
Wait(1)	
Call TakeScreenShot()

Call ActivateItemGuiTree("","Ship to other plant;Outbound Delivery","Outbound Delivery")
Call VerifyTextBoxContent("Outbound deliv\.","LIKP-VBELN","",DT_ME21N_1502_CHECK_TEXT_OF_OUTBOUND_DELIV,False)
Call ClickButton("Back   \(F3\)",False)

Call ActivateItemGuiTree("","Ship to other plant;Inbound Delivery","Inbound Delivery")
Call GetTextboxValue("LIKP-VBELN","","DT_ME21N_1602_CHECK_TEXT_OF_INBOUND_DELIV_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call ClickButton("Back   \(F3\)",False)

Call ActivateItemGuiTree("","Ship to other plant;Goods Issue","Goods Issue")
Call GetTextboxValue("GODYNPRO-MAT_DOC","","DT_ME21N_2010_CHECK_TEXT_OF_GODYNPROMAT_DOC_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call ClickButton("Back   \(F3\)",False)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()


