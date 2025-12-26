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

'.Mandatory Initial Call only in First Component in a Test Scenario
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name      : 
'.................Author : TCS          : Bitan
'................ Creation Date         : 7th April
'.................Modified By           :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Procurement of consumables from DC to Headquarter_p2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName=""


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet=2

'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  


'----------------------------------------VL10G---------------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()  
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Shipping Point/Receiving Pt","ST_VSTEL-LOW","",DT_VL10G_1000_SHIPPING_POINTRECEIVING_PT,False)
Call SetTextbox("Deliv\. Creation Date","ST_LEDAT-LOW","",ConvertDate(DT_VL10G_1000_DELIV_CREATION_DATE),False)
Call SetTextbox("to","ST_LEDAT-HIGH","",ConvertDate(DT_VL10G_1000_TO),False)
Call SetTextbox("CalcRuleDefltDlvCrDt","P_LERUL","","",False)
Call PressEnter()

Call SelectTab("TABSTRIP_ORDER_CRITERIA","Sales Orders",False)
'Call SetTextbox("Sales Document","ST_VBELN-LOW","",DT_VL10G_1020_SALES_DOCUMENT,False)
Call SetTextboxNoLabel("ST_VBELN-LOW","",DT_VL10G_1020_SALES_DOCUMENT,False)
Call TakeScreenShot
Call PressEnter()

Call SelectTab("TABSTRIP_ORDER_CRITERIA","Purchase Orders",False)
Call SetTextbox("Purchasing Document","ST_EBELN-LOW","",DT_VL10G_1030_PURCHASING_DOCUMENT,False)
Call PressEnter()     
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)

Call SelectRowGuiGridbyRowNo("",0,1,False)
Call ClickButtonIfExist("Create Delivery in Background   \(Shift\+F7\)",False)
Call TakeScreenShot

'Call GetGridContent("",0,"Sales Document",2,"Traffic light","S_TL_G","DT_VL10G_0500_OUTBOUND_DELIVERY_OUTPUT")

Call GetGridContent("",0,"SD Document",2,"Traffic light","S_TL_G","DT_VL10G_0500_OUTBOUND_DELIVERY_OUTPUT")

'--------------------------------/nME23N--------------------
Call SetTcode(DT_VL10G_0500_OKCD) 
Call PressEnter()
Call TakeScreenShot
'Call CheckTCodeScreen(DT_VL10G_0500_OKCD)

Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call SetTextbox("Pur. Order","MEPO_SELECT-EBELN","",DT_VL10G_1030_PURCHASING_DOCUMENT,True) 
Call SelectRadioButton("MEPO_SELECT-BSTYP_F","Pur. Order",True) 
Call ClickButton("Other Document   \(Enter\)",True)
Call TakeScreenShot

Call ClickButtonIfExist("Expand Item Details Ctrl\+F4",False)
Call SelectTab("ITEM_DETAIL","Purchase Order History",False)
Call VerifyGridCellContent("",1,"Article Document",0,DT_VL10G_0500_OUTBOUND_DELIVERY_OUTPUT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)


'----------------------------------/nVL03N------------------
Call SetTcode(DT_VL10G_4004_OKCD) 
Call PressEnter()
Call TakeScreenShot
'Call CheckTCodeScreen(DT_VL10G_4004_OKCD)

Call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_VL10G_4004_OUTBOUND_DELIVERY,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SelectMenuBar("Extras;Delivery Output;Header")
Call TakeScreenShot


Call VerifyTableCellContent(1, "Output Type", "SAPDV70ATC_NAST3",DT_VL10G_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_0)


'--------------------------------/nVL02N----------------------------
Call SetTcode(DT_VL10G_0100_OKCD) 
Call PressEnter()
Call TakeScreenShot
'Call CheckTCodeScreen(DT_VL10G_0100_OKCD)

Call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_VL10G_4004_OUTBOUND_DELIVERY_OCC1,False)
Call TakeScreenShot
Call ClickButton("Post Goods Issue   \(Shift\+F8\)",False)
Call TakeScreenShot
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
 ' VerifyStatusBarExist(Content)
Call VerifyStatusBarExist(DT_VL10G_4004_CHECK_TEXT_OF_STATUSBAR)


'--------------------------/nVL03n------------------------
Call SetTcode(DT_VL10G_4004_OKCD) 
Call PressEnter()
Call TakeScreenShot
'Call CheckTCodeScreen(DT_VL10G_4004_OKCD)

Call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_VL10G_4004_OUTBOUND_DELIVERY_OCC2,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call ClickButton("Document Flow   \(F7\)",False)
Call TakeScreenShot

Call ActivateNodeGuiTree("", "#1;#1;#1")
Call GetGridContent("TF.*", 0, "Doc.no.", 1, "<NA>", "<NA>", "DT_VL10G_0100_ARTICLE_DOCNUM_OUTPUT")

'
'--------------------------------------/nMIGO------------------------------
Call SetTcode(DT_VL10G_0100_OKCD_OCC1) 
Call PressEnter()
Call TakeScreenShot
'Call CheckTCodeScreen(DT_VL10G_0100_OKCD_OCC1)

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTextboxNoLabel("GODYNPRO-MAT_DOC", 0, DT_VL10G_0100_ARTICLE_DOCNUM,False)
Call ClickButtonIfExist("Continue   \(Enter\)", True)
Call PressEnter()
wait 2
Call PressEnter()

Call SelectTab("TS_GOITEM","Messages",False)
Call ClickButton("Display outputs",False)
Call VerifyTableCellContent(1,"Output Type","SAPDV70ATC_NAST3",DT_VL10G_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_0_OCC1)
Call ClickButton("Back   \(F3\)",False)

Call SelectTab("TS_GOHEAD","Doc. info",False)
Call ClickButton("FI Documents",False)
Call SelectRowGuiGridbyRowNo("Documents in Accounting","",1,True)
'Call SelectRowGuiGrid("Documents in Accounting", 0, "Object type text", "Accounting document", True)

Call GetGridContent("Documents in Accounting","","Document Number",1,"Object Type Text", "Accounting Document","DT_VL10G_0200_DOCNR")
Call ClickButtonIfExist("Display Document   \(F2\)",True)

Call VerifyGridCellContent("", 2, "Profit Center", 0,DT_VL10G_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)
Call VerifyGridCellContent("", 3, "Profit Center", 0, DT_VL10G_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_PRCTR)


Call LogOff()
Call FinalStatus ()

