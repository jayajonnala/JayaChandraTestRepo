

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_02ORD0202_005_Add_a_new_item_via_ME22N_or_WWP1
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

gstrTestCaseName = "Test_Add_a_new_item"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution1(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
''''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''
'---------------------------------  Tcode ME22N ------------------------------------------------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE)    
Call PressEnter() 
Call  CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)


call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call SetTextbox("Pur. Order","MEPO_SELECT-EBELN","",DT_ME22N_0003_PO,True)    
Call SelectRadioButton("MEPO_SELECT-BSTYP_F","Pur. Order",True)     
Call ClickButton("Other Document   \(Enter\)",True)     


Call ClickButtonIfExist("Expand Header Ctrl\+F2",False) 
Call GetspecialTextboxValue("MEPO_TOPLINE-SUPERFIELD","","DT_ME22N_1105_GET_TEXT_OF_VENDOR_OUTPUT",False)
Call SelectTab("HEADER_DETAIL","Org. Data",False)
If (VerifyTextBoxEnabled("Purch\. Org\.","MEPO1222-EKORG","",False)=False) Then
Call ClickButtonIfExist("Display/Change   \(F7\)",False)
End If
Call GetTextboxValue("MEPO1222-EKORG","","DT_ME22N_1221_GET_TEXT_OF_PURCH_ORG_OUTPUT",False)
Call GetTextboxValue("MEPO1222-EKGRP","","DT_ME22N_GET_CHECK_TEXT_OF_PURCH_GROUP_OUTPUT",False)
Call GetTextboxValue("MEPO1222-BUKRS","","DT_ME22N_1221_CGET_TEXT_OF_COMPANY_CODE_OUTPUT",False)
Call GetTextboxValue("MEPO_TOPLINE-BEDAT","","DT_ME22N_1221_CGET_TEXT_OF_DOC_DATE_OUTPUT",False)
'
Call GetTableCellData("SAPLMEGUITC_1211","Article",1,"","","DT_ME22N_1211_GET_TEXT_OF_TABLECELL_ARTICLE_0_OUTPUT",False)
Call GetTableCellData("SAPLMEGUITC_1211","Article",2,"","","DT_ME22N_1211_GET_TEXT_OF_TABLECELL_ARTICLE_1_OUTPUT",False)
Call GetTableCellData("SAPLMEGUITC_1211","Article",3,"","","DT_ME22N_1211_GET_TEXT_OF_TABLECELL_ARTICLE_2_OUTPUT",False)
Call GetTableCellData("SAPLMEGUITC_1211","Deliv. Date",1,"","","DT_ME22N_1211_GET_TEXT_OF_TABLECELL_DELIV_DATE_0_OUTPUT",False)
DELIV_DATE_0_OUTPUT= DT_ME22N_1211_GET_TEXT_OF_TABLECELL_DELIV_DATE_0_OUTPUT
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SelectRowGuiTable("SAPLMEGUITC_1211","Itm","10",False)
Call ClickButtonIfExist("Copy Item",False)
Call TakeScreenShot()
Call SelectMenuBar("Purchase Order;Save")
Call ClickButtonIfExist("Save",True)
Call VerifyifGuiLabelExists(DT_ME22N_120_CHECK_TEXT_SEVERAL_TIMES_ARTICLE)  ''--- Need to verify once message comes
Call ClickButton("Continue   \(Enter\)",True)    

Call ClickButtonIfExist("Collapse Item Details Ctrl\+F7",False)  
Call ClickButtonIfExist("DESELECT_ALL",False)
Call SelectRowGuiTable("SAPLMEGUITC_1211","Itm","20",False)
Call ClickButton("DELETE",False)
Call TakeScreenShot()
Call ClickButtonIfExist("Yes",True)
Call TakeScreenShot()

       
Call SetTableData("SAPLMEGUITC_1211","Article","2","","",DT_ME22N_1211_TABLECELL_ARTICLE_1,False)     
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","2","","",DT_ME22N_1211_TABLECELL_PO_QUANTITY_1,False) 
Call SetTableData("SAPLMEGUITC_1211","Plnt","2","","",DT_ME22N_1211_TABLECELL_SITE_1,False)   
Call PressEnter()
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","2","","",ConvertDate(DT_ME22N_1211_TABLECELL_DELIV_DATE_1),False)  
Call PressEnter()

Call SelectMenuBar("Purchase Order;Save")
Call ClickButtonIfExist("Save",True)
Call VerifyifGuiLabelExists(DT_ME22N_120_CHECK_TEXT_SEVERAL_DELIVERY_DATES)  ''--- Need to verify once message comes
Call ClickButton("Continue   \(Enter\)",True)    

Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","2","","",ConvertDate(DELIV_DATE_0_OUTPUT),False)  
Call PressEnter()

Call SelectMenuBar("Purchase Order;Save")
Call ClickButtonIfExist("Save",True)
Call VerifyifGuiLabelExists(DT_ME22N_120_CHECK_TEXT_MAINTAIN_SEVERAL_SITES)  ''--- Need to verify once message comes
Call ClickButton("Continue   \(Enter\)",True)    

Call SetTableData("SAPLMEGUITC_1211","Plnt","2","","",DT_ME22N_1211_TABLECELL_SITE_1_OCC4,False)   
Call PressEnter()
Call SelectMenuBar("Purchase Order;Save")
Call ClickButtonIfExist("Save",True)
Call VerifyStatusBarMessageType("S")
'
'
''
''---------------------------------  Tcode WWp1 ------------------------------------------------------------------
''
Call SetTcode(DT_ME22N_14_OKCD)    
Call PressEnter() 
Call  CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
Call SelectMenuBar("Goto;Variants;Get...")
Call SetTextbox("Created By","ENAME-LOW","","",True) 
Call SetTextbox("Original Language","MLANGU-LOW","","",True)
Call ClickButton("Execute   \(F8\)",True)     
Call SelectRowGuiGrid("Variant Catalog for Program RWWP001.*","","Variant name","AB",True)                  
Call PressEnter() 
'
Call SetTextbox("Site","01SWERKS-LOW","",DT_ME22N_0001_SITE,False)     
'Call SetTextbox("Vendor","01SLIFNR-LOW","",DT_ME22N_0001_VENDOR,False)   
Call SetTextboxNoLabel("01SLIFNR-LOW","",DT_ME22N_0001_VENDOR,False)
Call SetTextbox("Purchasing Group","01SSELEG-LOW","",DT_ME22N_0001_PURCHASING_GROUP,False)  
    
Call SetTextbox("Delivery date","01PEINDT","",DELIV_DATE_0_OUTPUT,False)  
''Call SetTextbox("Purchase Order Date","01PBEDAT","",DT_ME22N_001_DOC_DATE,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False) 
Call PressEnter() 
Call ClickButtonToolBar("SEA","")    
Call SetTextbox("Find:","G_STRING","",DT_ME22N_0003_PO,True)  
Call ClickButton("OK   \(Enter\)",True)   
Call TakeScreenShot()
Call ClickButtonToolBar("EXE","")  
Call  CheckScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2,"")
Call ClickButton("Additional Functions   \(Shift\+F9\)",False)  
Call  CheckScreen(DT_EXPECTEDTRANSACTIONCODE,"")
Call VerifyTableCellContent(2,"Itm","SAPLMEGUITC_1211","20")
 

Call LogOff()
Call FinalStatus ()





