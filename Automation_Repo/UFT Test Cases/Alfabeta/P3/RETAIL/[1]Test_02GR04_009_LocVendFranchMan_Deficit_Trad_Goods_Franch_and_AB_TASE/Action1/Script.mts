

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

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_02GR04_009_LocVendFranchMan_Deficit_Trad_Goods_Franch_and_AB
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 24th Oct
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_02GR04_009_LocVendFranchMan_Deficit_Trad_Goods_Franch_and_AB"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_02GR03_004_Loc_Vend_DSD_GR_Deliv_Note_w_Trading_Goods.xls"
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

'Enter Purchase Order Details with 3 items'
Call SetTableData("SAPLMEGUITC_1211","Article","1","","",DT_ME21N_1211_TABLECELL_ARTICLE_0,False) 
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_0,False)
Call SetTableData("SAPLMEGUITC_1211","OUn","1","","",DT_UNIT_OF_MESURE,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_ME21N_1211_TABLECELL_SITE_0,False) 
Call PressEnter()
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","1","","",ConvertDate(DT_DELIV_DATE),False) 
Call PressEnter()
Call SetTableData("SAPLMEGUITC_1211","Article","2","","",DT_ME21N_1211_TABLECELL_ARTICLE_1,False) 
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","2","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_1,False)
Call SetTableData("SAPLMEGUITC_1211","OUn","1","","",DT_UNIT_OF_MESURE,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","2","","",DT_ME21N_1211_TABLECELL_SITE_1,False) 
Call PressEnter()
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","2","","",ConvertDate(DT_DELIV_DATE),False) 
Call PressEnter()
Call SetTableData("SAPLMEGUITC_1211","Article","3","","",DT_ME21N_1211_TABLECELL_ARTICLE_2,False) 
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","3","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_2,False)
Call SetTableData("SAPLMEGUITC_1211","OUn","1","","",DT_UNIT_OF_MESURE,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","3","","",DT_ME21N_1211_TABLECELL_SITE_2,False) 
Call PressEnter()
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","3","","",ConvertDate(DT_DELIV_DATE),False) 
Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

'Click on Check Buton
Call SelectMenuBar("Purchase Order;Check")
Call TakeScreenShot()

'Click on Cancel
Call ClickButtonIfExist("Cancel   \(F12\)",True)
wait(2)

Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
'Navigate to Partner Tab
Call SelectTab("HEADER_DETAIL","Partners",False)
Wait(1)
Call TakeScreenShot()

'Enter details
Call SetTableData("SAPLEKPATC_0111","Funct","4","","",DT_ME21N_0111_TABLECELL_FUNCT_3,False) 
Call SetTableData("SAPLEKPATC_0111","Number","4","","",DT_ME21N_0111_TABLECELL_NUMBER_3,False) 
Call TakeScreenShot()

'Click on Check Buton
Call SelectMenuBar("Purchase Order;Check")

Call VerifyStatusBarMessageType("S")

'Click on Save
Call SelectMenuBar("Purchase Order;Save")

Call VerifyStatusBarMessageType("S")

'Validate If Purchase order is generated
Call GetStatusBar("item2","DT_TRIANGULAR_SALESORDER_OUTPUT")
VerifyStatusBar("AB Triangular Sales created under the number " & DT_TRIANGULAR_SALESORDER_OUTPUT)

'----------------------Tcode MIGO----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_ME21N_14_OKCD) 
Call PressEnter() 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
Wait(2)

'Enter PO Details
Call SetCombo("GODYNPRO-ACTION","Goods Receipt")
Call SetCombo("GODYNPRO-REFDOC","Purchase Order")
Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER",0,DT_ME21N_2000_GODYNPROPO_NUMBER,False)
Call SetTextboxNoLabel("GODEFAULT_TV-BWART",0,DT_ME21N_0010_GODEFAULT_TVBWART,False)
Wait(2)

'Set Delivey Note
Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_ME21N_0110_DELIVERY_NOTE,False)
Call ClickButtonIfExist("MIGO_OK_GO",False)
Call SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,"ON",False)
Call TakeScreenShot()

'Click on Check Button
Call ClickButton("Check Entries   \(F7\)",False)
Wait(5)

Call VerifyStatusBarMessageType("S")

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call VerifyStatusBarMessageType("S") 
Call GetStatusBar("item1","DT_ARTICLE_NO_OUTPUT")
Call VerifyStatusBar("Article document " & DT_ARTICLE_NO_OUTPUT&" "&"posted")

''------------------------------Display Article Document Details-------------------------------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetCombo("GODYNPRO-ACTION","Display")
Call SetComboByKey("GODYNPRO-REFDOC","R02")
'Call SetCombo("GODYNPRO-REFDOC","Article Document")
'Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER",0,DT_ARTICLE_DOCUMENT,False)
'Call SetTextboxNoLabel("GODYNPRO-DOC_YEAR",0,2021,False)
Call PressEnter() 
Wait(2)
Call TakeScreenShot()

'Navigate to the Document Details
Call SelectTab("TS_GOHEAD","Doc. info",False)
Wait(1)
Call TakeScreenShot()
Call ClickButton("FI Documents",False)
Wait(5)

'Check Screen
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)

'Get Document Number
Call GetTextboxValue("BKPF-BELNR",0,"DT_FB03_DOCUMENT_NUM",False)

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

'Check Screen
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)

''----------------------Tcode WE09----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_ME21N_1_OKCD) 
Call PressEnter() 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC5)
Wait(2)

'Fill The Details and Execute
Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call ClickButtonIfExist("Execute   \(F8\)",True) 
Call SelectRowGuiGrid("Variant Catalog for Program RSEIDOC9.*","","Variant name","/MBGMCR_OUT",True)
Call ClickButtonIfExist("Choose   \(F2\)",True)
Call SetTextboxNoLabel("VALUE1_1",0,DT_ME21N_1000_FOR_VALUE,False)
Wait(2)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False) 
Wait(2)
Call ClickButtonIfExist("Yes",True)
Wait(5)
Call VerifyStatusBar("IDocs were found")

'Navigate to Idoc Tree and Verify the details
'Call GetLabelContentByRefLabel(Right(DT_ME21N_100_VARIANT,Len(DT_ME21N_100_VARIANT)),994,0,"DT_IDOC_NUMBER",False)
Call GetLabelContentByRefLabel(Right(DT_ME21N_100_VARIANT,Len(DT_ME21N_100_VARIANT)),1136,0,"DT_IDOC_NUMBER",False)
Call ClickLabel(DT_IDOC_NUMBER,0,False)
Call ClickLinkGuiTree("IDoc "& DT_IDOC_NUMBER&";Data records;#1;E1BP2017_GM_HEAD_01","E1BP2017_GM_HEAD_01","",False)
Call VerifyTableCellContent(3,"Fld Cont.","IDOC_TREE_CONTROLINT_SEG_CONTROL",DT_ME21N_0110_DELIVERY_NOTE)
Call VerifyTableCellContent(7,"Fld Cont.","IDOC_TREE_CONTROLINT_SEG_CONTROL",DT_ME21N_1000_FOR_VALUE)
Call ClickLinkGuiTree("IDoc "& DT_IDOC_NUMBER&";Data records;#1;E1BP2017_GM_ITEM_CREATE","E1BP2017_GM_ITEM_CREATE","",False)
Call TakeScreenShot()
Call ClickLinkGuiTree("IDoc "& DT_IDOC_NUMBER&";Data records;E1MBGMCR;#2;E1BP2017_GM_ITEM_CREATE1","E1BP2017_GM_ITEM_CREATE1","",False)
Call TakeScreenShot()
Call VerifyTableCellContent(8,"Fld Cont.","IDOC_TREE_CONTROLINT_SEG_CONTROL",DT_ME21N_100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_7)

Call ClickLinkGuiTree("IDoc "& DT_IDOC_NUMBER&";Data records;E1MBGMCR;#2;E1BP2017_GM_ITEM_CREATE1","E1BP2017_GM_ITEM_CREATE1","",False)
Call TakeScreenShot()
Call VerifyTableCellContent(8,"Fld Cont.","IDOC_TREE_CONTROLINT_SEG_CONTROL",DT_ME21N_100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_7)


Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
