
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_159_Update Account Completion Table (All Opcos)_TASE
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)


gstrTestCaseName = "Test_06-01-02-02-02-Check MSI Net Orders in SAP_PO"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
Call ClickButtonIfExist("Continue (Enter)",True)

'-----------------------------------------ME23N---------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

''''Added Line 49 to 52' On 15Mar 2022 for AT_BE_Interco Dry
''' Author - KGARA
'Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
'Call SetTextbox("Pur. Order","MEPO_SELECT-EBELN","",DT_SAPOPTIONS,True)    
'Call SelectRadioButton("MEPO_SELECT-BSTYP_F","Pur. Order",True)     
'Call ClickButton("Other Document   \(Enter\)",True)  

Call ClickButtonIfExist("Switch On Document Overview   \(F8\)",false)  
call ClickContextButtonToolBar("SELECT",0)
Call TakeScreenShot()
call SelectMenuItemToolBar("Purchase orders",0)
Call ClickButton("Dynamic selections   \(Shift\+F4\)",False) 
Call TakeScreenShot()

Call ActivateNodeGuiTree(0,"Purchasing Document Header;Your Reference")

Call SetTextbox("Your Reference","%%DYN001-LOW","",DT_ME23N_1106_YOUR_REFERENCE,false)    
Call SetTextbox("Purchasing Document","SP\$00014-LOW","","",false)  
Call SetTextbox("Document Date","SP\$00011-LOW","",ConvertDate(DT_DOC_DATE),False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)  
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot()
call ActivateNodeGuiTree(0,"#1:5522776211")

Call TakeScreenShot()
Call ClickButtonIfExist("Expand Item Details Ctrl\+F4",False)
Call SelectTab("ITEM_DETAIL","Quantities/Weights",False)
call ClickButtonIfExist("Expand Items Ctrl\+F3",false)

Call TakeScreenShot()

Call VerifyTableCellContent(1,"Article","SAPLMEGUITC_1211",DT_ME23N_1211_CHECK_TEXT_OF_TABLECELL_ARTICLE_0)
Call VerifyTableCellContent(2,"Article","SAPLMEGUITC_1211",DT_ME23N_1211_CHECK_TEXT_OF_TABLECELL_ARTICLE_1)
Call VerifyTableCellContent(1,"PO Quantity","SAPLMEGUITC_1211",DT_ME23N_1211_CHECK_TEXT_OF_TABLECELL_PO_QUANTITY_0)
Call VerifyTableCellContent(2,"PO Quantity","SAPLMEGUITC_1211",DT_ME23N_1211_CHECK_TEXT_OF_TABLECELL_PO_QUANTITY_1)
Call VerifyTextBoxNoLabelContent("MEPO_TOPLINE-SUPERFIELD","",DT_ME23N_1105_CHECK_TEXT_OF_SUPPLYING_SITE,False)
Call VerifyTableCellContent(1,"Plnt","SAPLMEGUITC_1211",DT_SITE_1)
Call VerifyTableCellContent(2,"Plnt","SAPLMEGUITC_1211",DT_SITE_2)
'' VerifyComboBoxValue(comboAttachedText, expectedValue)
''Call VerifyComboBoxValue("",DT_ME23N_1105_CHECK_VALUE_OF_MEPO_TOPLINEBSART)

call GetTextboxValue("MEPO_TOPLINE-EBELN",1,"Article_num",false)
call WriteRunTimeDataToExcel ("DT_ME23N_1105_GET_TEXT_OF_MEPO_TOPLINEEBELN_Output",Article_num)
''
''Call VerifyTableCellContent(1,"Deliv. Date","SAPLMEGUITC_1211",DT_DOC_DATE)'''''DT_ME23N_1211_CHECK_TEXT_OF_TABLECELL_DELIV_DATE_0)
''Call VerifyTableCellContent(2,"Deliv. Date","SAPLMEGUITC_1211",DT_DOC_DATE)''''DT_ME23N_1211_CHECK_TEXT_OF_TABLECELL_DELIV_DATE_1)
''
call ClickButtonIfExist("Expand Header Ctrl\+F2",false)
Call SelectTab("HEADER_DETAIL","Org. Data",False) 
Call TakeScreenShot()
call VerifyTextBoxContent("Purch\. Org\.","MEPO1222-EKORG",1,DT_ME23N_1221_CHECK_TEXT_OF_PURCH_ORG,false)
call VerifyTextBoxContent("Purch\. Group","MEPO1222-EKGRP",1,DT_ME23N_1221_CHECK_TEXT_OF_PURCH_GROUP,false)
call VerifyTextBoxContent("Company Code","MEPO1222-BUKRS",1,DT_ME23N_1221_CHECK_TEXT_OF_COMPANY_CODE,false)
Call TakeScreenShot()
'Call SetTcode(DT_ME23N_0014_OKCD)    

Call LogOff()
Call FinalStatus ()

''
'Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
'Call SetTextbox("Pur. Order","MEPO_SELECT-EBELN","",DT_SAPOPTIONS,True)    
'Call SelectRadioButton("MEPO_SELECT-BSTYP_F","Pur. Order",True)     
'Call ClickButton("Other Document   \(Enter\)",True)  
'
''Call ClickButton("Switch On Document Overview   \(F8\)",false)  
'
'
'' ClickButtonToolBar(buttonName, toolbarIndex)
''call ClickButtonToolBar("shell\[0\]",0)
'' SelectMenuItemToolBar(menuItem, toolbarIndex)
'
'
'' ClickContextButtonToolBar(buttonName, toolbarIndex)
''call ClickContextButtonToolBar("SELECT",0)
'
'' Click204ButtonToolBar(buttonName, toolbarIndex)
''Click204ButtonToolBar
'' ClickButtonToolBar(buttonName, toolbarIndex)
''ClickButtonToolBar
'' ClickButtonIfExist(tooltipOrButtonName, blnIsItPopup)
''call ClickButtonIfExist("shell",false)
''Call ClickButtonToolBar("shell", 4)
''
''SAPGuiSession("Session").SAPGuiWindow("Standard PO Retail 4304284516").Maximize
''SAPGuiSession("Session").SAPGuiWindow("Standard PO Retail 4304284516").SAPGuiTable("SAPLMEGUITC_1211").ClickCell 3,"Status"
''SAPGuiSession("Session").SAPGuiWindow("Standard PO Retail 4304284516").SAPGuiToolbar("ToolBarControl").PressButton "HIDEHELP"
'
'' SelectMenuItemToolBar(menuItem, toolbarIndex)
'
''SAPGuiSession("Session_2").SAPGuiWindow("STO Retail 5535527521").Maximize
''SAPGuiSession("Session_2").SAPGuiWindow("STO Retail 5535527521").Page("Page").Frame("Frame0").WebElement("No variant defined").Click
''SAPGuiSession("Session_2").SAPGuiWindow("STO Retail 5535527521").SAPGuiToolbar("ToolBarControl").PressContextButton "SELECT"
''SAPGuiSession("Session_2").SAPGuiWindow("STO Retail 5535527521").SAPGuiToolbar("ToolBarControl").SelectMenuItem "Purchase orders"
''call ClickButtonToolBar("SELECT",False)
''call SelectMenuItemToolBar("Purchase orders",1)
'
'call ClickContextButtonToolBar("SELECT",0)
'Call TakeScreenShot()
'call SelectMenuItemToolBar("Purchase orders",0)
'Call ClickButton("Dynamic selections   \(Shift\+F4\)",False) 
'Call TakeScreenShot()
'
'Call SelectNodeGuiTree(0,"Purchasing Document Header;Your Reference")
' 
'call ClickButtonToolBar("TAKE",0)
'Call SetTextbox("Your Reference","%%DYN001-LOW","",DT_ME23N_1106_YOUR_REFERENCE,false)    
'Call SetTextbox("Purchasing Document","SP\$00014-LOW","","",false)    
'
'
'Call ClickButton("Execute   \(F8\)",False)  
'Call TakeScreenShot()
''Added Line 95 to 100 ' On 15Mar 2022
'' Author - KGARA
'Call ClickButtonIfExist("Continue   \(Enter\)",True)
'call ClickButtonToolBar("shell\[0\]",0)
'Call ClickButtonIfExist("Execute   \(F8\)",True)
'Call ClickButtonIfExist("Continue   \(Enter\)",True)
'' ActivateSAPUITree(creationTime, windowText, htmlTag, htmlId, innertext, className, index, title, path, blnIsItPopup)
''ActivateSAPUITree
'' ActivateNodeGuiTree(treeIndex, itemPath)
'call ActivateNodeGuiTree(0,"#1:5522776211")
'' VerifyTableCellContent(tableRowNumber, tableColumnName, tableName, expectedValue)
'
''Added Line 111 ' On 15Mar 2022
'' Author - KGARA
'
'Call SelectTab("ITEM_DETAIL","Quantities/Weights",False)
'
'call VerifyTableCellContent(1,"Article","SAPLMEGUITC_1211",DT_ME23N_1211_CHECK_TEXT_OF_TABLECELL_ARTICLE_0)
'call VerifyTableCellContent(2,"Article","SAPLMEGUITC_1211",DT_ME23N_1211_CHECK_TEXT_OF_TABLECELL_ARTICLE_1)
'call VerifyTableCellContent(1,"PO Quantity","SAPLMEGUITC_1211",DT_ME23N_1211_CHECK_TEXT_OF_TABLECELL_PO_QUANTITY_0)
'call VerifyTableCellContent(2,"PO Quantity","SAPLMEGUITC_1211",DT_ME23N_1211_CHECK_TEXT_OF_TABLECELL_PO_QUANTITY_1)
''call VerifyTextBoxContent("Information Message","GS_SEARCH-SEARCH_INFO",1,DT_ME23N_0841_CHECK_TEXT_OF_GS_SEARCHSEARCH_INFO,FALSE)
'
'' VerifyTextBoxContent(textboxAttachedText, textboxName, textboxIndex, expectedValue, blnIsI
'
''call VerifyTextBoxContent("Supplying Site","MEPO_TOPLINE-SUPERFIELD",1,DT_ME23N_1105_CHECK_TEXT_OF_SUPPLYING_SITE,false)
''call VerifyTextBoxContent(".*","MEPO_TOPLINE-SUPERFIELD",1,DT_ME23N_1105_CHECK_TEXT_OF_SUPPLYING_SITE,false)
'call GetTextboxValue("MEPO_TOPLINE-EBELN",1,"Article_num",false)
'call WriteRunTimeDataToExcel ("DT_ME23N_1105_GET_TEXT_OF_MEPO_TOPLINEEBELN_Output",Article_num)
'
''call VerifyTableCellContent(1,"Deliv.Date","SAPLMEGUITC_1211",DT_ME23N_1211_CHECK_TEXT_OF_TABLECELL_DELIV_DATE_0)
''call VerifyTableCellContent(2,"Deliv.Date","SAPLMEGUITC_1211",DT_ME23N_1211_CHECK_TEXT_OF_TABLECELL_DELIV_DATE_1)
''call VerifyTextBoxContent(".*","MEPO_TOPLINE-SUPERFIELD",1,DT_ME23N_1105_CHECK_TEXT_OF_SUPPLYING_SITE,false)
'call ClickButtonIfExist("Expand Header Ctrl\+F2",false)
''Call SelectTab("HEADER_DETAIL","Org\. Data",False) 
'Call SelectTab("HEADER_DETAIL","Org. Data",False) 
'call VerifyTextBoxContent("Purch\. Org\.","MEPO1222-EKORG",1,DT_ME23N_1221_CHECK_TEXT_OF_PURCH_ORG,false)
'call VerifyTextBoxContent("Purch\. Group","MEPO1222-EKGRP",1,DT_ME23N_1221_CHECK_TEXT_OF_PURCH_GROUP,false)
'call VerifyTextBoxContent("Company Code","MEPO1222-BUKRS",1,DT_ME23N_1221_CHECK_TEXT_OF_COMPANY_CODE,false)
'Call SetTcode(DT_ME23N_0014_OKCD)    
'
'
'Call LogOff()
'Call FinalStatus ()

