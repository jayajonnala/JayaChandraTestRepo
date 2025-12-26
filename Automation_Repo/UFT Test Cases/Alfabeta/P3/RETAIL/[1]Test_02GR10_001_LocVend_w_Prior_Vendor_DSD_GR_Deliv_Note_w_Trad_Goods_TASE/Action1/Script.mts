
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_02GR10_001_LocVend_w_Prior_Vendor_DSD_GR_Deliv_Note_w_Trad_Goods
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 2nd Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_02GR10_001_LocVend_w_Prior_Vendor_DSD_GR_Deliv_Note_w_Trad_Goods"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_02GR10_001_LocVend_w_Prior_Vendor_DSD_GR_Deliv_Note_w_Trad_Goods.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode ME21N----------------------------

'Create Purchase Order
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

'Enter the Details
Call SetComboByKey("MEPO_TOPLINE-BSART",DT_ME21N_1105_MEPO_TOPLINEBSART)
'Call SetTextbox("Vendor","MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,False)   
Call SetTextboxNoLabel("MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,False)
Call SetTextbox("Purch. Org.","MEPO1222-EKORG","",DT_ME21N_1221_PURCH_ORG,False) 
Call SetTextbox("Purch. Group","MEPO1222-EKGRP","",DT_ME21N_1221_PURCH_GROUP,False)     
Call PressEnter() 
Call SetTextbox("Company Code","MEPO1222-BUKRS","",DT_ME21N_1221_COMPANY_CODE,False)    
Call PressEnter()

'Enter Purchase Order Details'
Call SetTableData("SAPLMEGUITC_1211","Article","1","","",DT_ME21N_1211_TABLECELL_ARTICLE_0,False) 
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_0,False)

Call SetTableData("SAPLMEGUITC_1211","OUn","1","","",DT_ME21N_1211_TABLECELL_OUN_0,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_ME21N_1211_TABLECELL_SITE_0,False) 
Call PressEnter()
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","1","","",DT_ME21N_1211_TABLECELL_DELIV_DATE_0,False) 
Call PressEnter()

'Enter Purchase Order Details'
Call SetTableData("SAPLMEGUITC_1211","Article","2","","",DT_ME21N_1211_TABLECELL_ARTICLE_1,False) 
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","2","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_1,False)

Call SetTableData("SAPLMEGUITC_1211","OUn","2","","",DT_ME21N_1211_TABLECELL_OUN_1,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","2","","",DT_ME21N_1211_TABLECELL_SITE_1,False) 
Call PressEnter()
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","2","","",DT_ME21N_1211_TABLECELL_DELIV_DATE_1,False) 
Call PressEnter()

'Capture the screenshot
Call TakeScreenShot()

'Click on Save Buton
Call SelectMenuBar("Purchase Order;Save")
Call ClickButtonIfExist("Save",True)
wait(2)

'Validate If Purchase order is generated
Call GetStatusBar("item2","DT_PO_NUMBER_OUTPUT")
VerifyStatusBar("Direct Delivery created under the number " & DT_PO_NUMBER_OUTPUT)
'----------------------Tcode MIGO----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_ME21N_14_OKCD) 
Call PressEnter() 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
Wait(2)
Call SetCombo("GODYNPRO-ACTION","Display")
'''Call SetCombo("GODYNPRO-REFDOC","Article Document")
Wait(2)
Call TakeScreenShot()

Call SetCombo("GODYNPRO-ACTION","Goods Receipt")
Call SetCombo("GODYNPRO-REFDOC","Purchase Order")
Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER",0,DT_PO_NUMBER_OUTPUT,False)
Wait(2)
Call SetTextboxNoLabel("GODEFAULT_TV-BWART",0,DT_ME21N_0010_GODEFAULT_TVBWART_OCC2,False)
Wait(2)

Call SelectTab("TS_GOHEAD","General",False)
Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_ME21N_0110_DELIVERY_NOTE,False)
Call ClickButtonIfExist("MIGO_OK_GO",False)

Call ClickButtonIfExist("Open detail data",False)
Call SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,"ON",False)
Call TakeScreenShot()

Call SetTextbox("Line","GODYNPRO-DETAIL_ZEILE","",2,False)
Call ClickButtonIfExist("OK_LOCATE",False)
Call TakeScreenShot()

Call ClickButton("Check Entries   \(F7\)",False)
Wait(1)
'VerifyStatusBar(DT_MIGO_1_CHECK_TEXT_OF_STATUSBAR)

'Post the Document
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True) 
Wait(2)
Call GetStatusBar("item1","DT_ARTICLE_NO_OUTPUT")
VerifyStatusBar("Article document " & DT_ARTICLE_NO_OUTPUT&" "&"posted")

''------------------------------Display Article Document Details-------------------------------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetCombo("GODYNPRO-ACTION","Display")
'Call SetCombo("GODYNPRO-REFDOC","Article Document")
'Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER",0,DT_ARTICLE_NO_OUTPUT,False)
'Call SetTextboxNoLabel("GODYNPRO-DOC_YEAR",0,2020,False)
Call PressEnter() 
Wait(2)
Call TakeScreenShot()

'Navigate to the Document Details
Call SelectTab("TS_GOHEAD","Doc. info",False)
Wait(1)
Call TakeScreenShot()
Call ClickButton("FI Documents",False)
Wait(5)
Call TakeScreenShot()

'Verify If FB03 Screen is displayed
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)



'''------------------------------TCode WE9-------------------------------------------------

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_ME21N_1_OKCD) 
Call PressEnter() 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC5)
'
'Fill The Details and Execute
Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call ClickButtonIfExist("Execute   \(F8\)",True) 
Call SelectRowGuiGrid("Variant Catalog for Program RSEIDOC9.*","","Variant name","/MBGMCR_OUT",True)
Call ClickButtonIfExist("Choose   \(F2\)",True) 
Wait(2)
Call SetTextboxNoLabel("VALUE1_1",0,DT_ME21N_1000_FOR_VALUE__OCC2,False)
Wait(3)
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False) 
Wait(2)
Call ClickButtonIfExist("Yes",True)
Wait(5)
Call VerifyStatusBar("IDocs were found")

Call GetLabelContentByRefLabel(Right(DT_ME21N_100_VARIANT,Len(DT_ME21N_100_VARIANT)),1136,0,"DT_IDOC_NUMBER",False)

if VerifyifGuiLabelCondition(DT_MIGO_120_CHECK_TEXT_OF_DATA_PASSED_TO_PORT_OK) = True Then

	Call ClickLabel(DT_IDOC_NUMBER,0,False)
	Call ClickLinkGuiTree("IDoc "& DT_IDOC_NUMBER&";Data records;#1;E1BP2017_GM_HEAD_01","E1BP2017_GM_HEAD_01","",False)
	Call VerifyTableCellContent(7,"Fld Cont.","IDOC_TREE_CONTROLINT_SEG_CONTROL",DT_ME21N_1000_FOR_VALUE__OCC2)

Else

	Call SetTcode(DT_BD87_OKCD) 
	Call PressEnter() 
	Call TakeScreenshot()
	Call SetTextbox("IDoc Number","SX_DOCNU-LOW","",DT_IDOC_NUMBER,False)
	Call TakeScreenShot()
	Call ClickButton("Execute   \(F8\)",False)
	Call TakeScreenShot()
	Call SelectItemGuiTree(0,"Retail Pre-Production;IDocs in outbound processing;IDoc ready for dispatch (ALE service)","IDoc ready for dispatch (ALE service)")
	Call TakeScreenShot()
	Call ClickButton("Process Selected Node   \(F8\)",False)
	Call TakeScreenShot()
	Call ClickButton("Continue   \(Enter\)",True)
	Call TakeScreenShot()
	Call VerifyGridCellContent("Processed IDocs", 1, "Status text", 0, DT_ME21N_120_CHECK_TEXT_OF_DATA_PASSED_TO_PORT_OK)

	Call SetTcode(DT_ME21N_1_OKCD) 
	Call PressEnter() 
	Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC5)
	'
	'Fill The Details and Execute
	Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
	Call SetTextbox("Created By","ENAME-LOW","","",True)
	Call ClickButtonIfExist("Execute   \(F8\)",True) 
	Call SelectRowGuiGrid("Variant Catalog for Program RSEIDOC9.*","","Variant name","/MBGMCR_OUT",True)
	Call ClickButtonIfExist("Choose   \(F2\)",True) 
	Wait(2)
	Call SetTextboxNoLabel("VALUE1_1",0,DT_ME21N_1000_FOR_VALUE__OCC2,False)
	Wait(3)
	Call TakeScreenShot()
	
	Call ClickButton("Execute   \(F8\)",False) 
	Wait(2)
	Call ClickButtonIfExist("Yes",True)
	Wait(5)
	Call VerifyStatusBar("IDocs were found")

End if




'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

