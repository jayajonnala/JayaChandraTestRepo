

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


'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_02GR03_007_Loc_Vend_DSD_GR_Deliv_Note_w_free_goods
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 24th Oct
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_02GR03_007_Loc_Vend_DSD_GR_Deliv_Note_w_free_goods"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_02GR03_007_Loc_Vend_DSD_GR_Deliv_Note_w_free_goods.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
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
Call TakeScreenShot()

'Enter the Details
Call SetComboByKey("MEPO_TOPLINE-BSART",DT_ME21N_1105_MEPO_TOPLINEBSART)
'Call SetTextbox("Vendor","MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,False)   
Call SetTextboxNoLabel("MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,False)
Call SetTextbox("Purch. Org.","MEPO1222-EKORG","",DT_ME21N_1221_PURCH_ORG,False) 
Call SetTextbox("Purch. Group","MEPO1222-EKGRP","",DT_ME21N_1221_PURCH_GROUP,False)     
Call PressEnter() 
Call SetTextbox("Company Code","MEPO1222-BUKRS","",DT_ME21N_1221_COMPANY_CODE,False)    
Call PressEnter()
Call TakeScreenShot()

'Enter Purchase Order Details with 3 items'
Call SetTableData("SAPLMEGUITC_1211","Article","1","","",DT_ME21N_1211_TABLECELL_ARTICLE_0,False) 
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_0,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_ME21N_1211_TABLECELL_SITE_0,False) 
Call PressEnter()
Wait(1)
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","1","","",ConvertDate(DT_ME21N_1211_TABLECELL_DELIV_DATE_0),False) 
Call PressEnter()
Call SetTableData("SAPLMEGUITC_1211","Article","2","","",DT_ME21N_1211_TABLECELL_ARTICLE_1,False) 
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","2","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_1,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","2","","",DT_ME21N_1211_TABLECELL_SITE_1,False) 
Call PressEnter()
Wait(1)
Call TakeScreenShot()
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","2","","",ConvertDate(DT_ME21N_1211_TABLECELL_DELIV_DATE_1),False) 
Call PressEnter()
Call SetTableData("SAPLMEGUITC_1211","Article","3","","",DT_ME21N_1211_TABLECELL_ARTICLE_2,False) 
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","3","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_2,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","3","","",DT_ME21N_1211_TABLECELL_SITE_2,False) 
Call PressEnter()
Wait(1)
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","3","","",ConvertDate(DT_ME21N_1211_TABLECELL_DELIV_DATE_2),False) 
Call PressEnter()


'Capture the screenshot
Call TakeScreenShot()

'Click on Save Buton
Call SelectMenuBar("Purchase Order;Save")
Call ClickButtonIfExist("Save",True)
wait(2)
Call TakeScreenShot()
'Validate If Purchase order is generated
Call GetStatusBar("item2","DT_PO_NUMBER_OUTPUT")
VerifyStatusBar("Direct Delivery created under the number " & DT_PO_NUMBER_OUTPUT)
''Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
''Call PressEnter()
'Call SetComboByKey("DYN_6000-LIST","   3")
'Call SelectTab("ITEM_DETAIL","Condition Control",False)
'Call TakeScreenShot()
''''----------------------Tcode MIGO----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_ME21N_14_OKCD) 
Call PressEnter() 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
Call TakeScreenShot()
Wait(2)
Call SetCombo("GODYNPRO-ACTION","Goods Receipt")
Call SetCombo("GODYNPRO-REFDOC","Purchase Order")
Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER",0,DT_ME21N_2000_GODYNPROPO_NUMBER,False)
Call SetTextboxNoLabel("GODEFAULT_TV-BWART",0,DT_MIGO_0010_GODEFAULTTV_BWART,False)
Wait(2)
Call TakeScreenShot()
Call WriteRunTimeDataToExcelGlobalSheet ("DT_NUM",Cint(DT_NUM)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SelectTab("TS_GOHEAD","General",False)
Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_ME21N_0110_DELIVERY_NOTE,False)
Call ClickButtonIfExist("MIGO_OK_GO",False)
Call ClickButtonIfExist("Open detail data",False)
Call TakeScreenShot()
Call SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,"ON",False)
Call TakeScreenShot()

Call SetTextbox("Line","GODYNPRO-DETAIL_ZEILE","",2,False)
Call ClickButtonIfExist("OK_LOCATE",False)
Call TakeScreenShot()

Call SetTextbox("Line","GODYNPRO-DETAIL_ZEILE","",3,False)
Call ClickButtonIfExist("OK_LOCATE",False)
Call TakeScreenShot()
Call ClickButton("Check Entries   \(F7\)",False)
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  
Call TakeScreenShot()
Call GetStatusBar("item1","DT_ARTICLE_NO_OUTPUT")
VerifyStatusBar("Article document " & DT_ARTICLE_NO_OUTPUT&" "&"posted")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''------------------------------Display Article Document Details-------------------------------------------------

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

''''------------------------------TCode WE09-------------------------------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_ME21N_750_OKCD) 
Call PressEnter() 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)
'
'Fill The Details and Execute
Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call ClickButtonIfExist("Execute   \(F8\)",True) 
Call SelectRowGuiGrid("Variant Catalog for Program RSEIDOC9.*","","Variant name","/MBGMCR_OUT",True)
Call ClickButtonIfExist("Choose   \(F2\)",True) 
Call SetTextboxNoLabel("VALUE1_1",0,DT_ME21N_1000_FOR_VALUE__OCC2,False)
Wait(3)

'Call SetTextbox("for Value \.\.\.","VALUE1_1","",DT_ME21N_1000_FOR_VALUE__OCC2,False)
'Wait(3)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False) 
Wait(2)
Call ClickButtonIfExist("Yes",True)
Wait(5)
Call VerifyStatusBar("IDocs were found")

Call GetLabelContentByRefLabel(Right(DT_ME21N_600_GRIDCELL_3_VARIANT_NAME_OCC2,Len(DT_ME21N_600_GRIDCELL_3_VARIANT_NAME_OCC2)),994,0,"DT_IDOC_NUMBER",False)
'Call VerifyifGuiLabelExists(DT_ME21N_120_CHECK_TEXT_OF_DATA_PASSED_TO_PORT_OK)
Call ClickLabel(DT_IDOC_NUMBER,0,False)
Call ClickLinkGuiTree("IDoc "& DT_IDOC_NUMBER&";Data records;#1;E1BP2017_GM_HEAD_01","E1BP2017_GM_HEAD_01","",False)
'Call VerifyTableCellContent(6,"Fld Cont.","IDOC_TREE_CONTROLINT_SEG_CONTROL",DT_ME21N_100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_6)
Call ClickLinkGuiTree("IDoc "& DT_IDOC_NUMBER&";Data records;#1;E1BP2017_GM_ITEM_CREATE","E1BP2017_GM_ITEM_CREATE","",False)
Call TakeScreenShot()


'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

