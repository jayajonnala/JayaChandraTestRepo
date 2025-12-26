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

  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'reload DS to update dates and calculations
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_S2A_PRI_02_049-Compare PBandNB prices_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 15th June
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_S2A_PRI_02_049-Compare PBandNB prices_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_S2A_PRI_02_049-Compare PBandNB prices_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'''Increment the parameter/reload
''Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call WriteRunTimeDataToExcelGlobalSheet ("DT_XYZ",1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''
''''----------------------Tcode VK11----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Condition Type","RV13A-KSCHL","",DT_VK11_0100_CONDITION_TYPE,False)
Call ClickButtonIfExist("Key Combination   \(Shift\+F5\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Choose   \(Enter\)",True)

Call SetTextbox("Sales Organization","KOMG-VKORG","",DT_VK11_1073_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution Channel","KOMG-VTWEG","",DT_VK11_1073_DISTRIBUTION_CHANNEL,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Article",1,DT_VK11_1073_TABLECELL_ARTICLE_0,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Sales unit",1,DT_VK11_1073_TABLECELL_SALES_UNIT_0,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Amount",1,DT_VK11_1073_TABLECELL_AMOUNT_0,False)
Call PressEnter() 

Call GetTableCellData("SAPMV13ATCTRL_FAST_ENTRY","Amount",1,"Article",DT_VK11_1073_TABLECELL_ARTICLE_0,"DT_VK11_1073_GET_TABLECELL_AMOUNT_0_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SelectRowGuiTableByRow("SAPMV13ATCTRL_FAST_ENTRY",1,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Details   \(F6\)",False)

Call SetTextbox("Lower limit","KONP-MXWRT","",DT_VK11_0300_LOWER_LIMIT,False)
Call SetTextbox("Upper limit","KONP-GKWRT","",DT_VK11_0300_UPPER_LIMIT,False)
Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()
Call ClickButtonIfExist("Back   \(F3\)",False)

'Capture the screenshot
Call TakeScreenShot()
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()
Call VerifyStatusBar(DT_VK11_1073_CHECK_TEXT_OF_STATUSBAR)
'
''
''''----------------------Tcode VK13----------------------------
'Enter the Tcode
Call SetTcode(DT_VK11_1073_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_VK11_1073_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Condition Type","RV13A-KSCHL","",DT_VK11_0100_CONDITION_TYPE_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButtonIfExist("Condition Information   \(Shift\+F4\)",False)

Call SetTextbox("Sales Organization","F001-LOW","",DT_VK11_1000_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution Channel","F002-LOW","",DT_VK11_1000_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Article","F003-LOW","",DT_VK11_1000_ARTICLE,False)
Call SetTextbox("from / on","SEL_DATE","",DT_VK11_1000_FROM__ON,False)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButtonIfExist("Execute   \(F8\)",False)

Call SelectCheckboxNoLabel("1",DT_VK11_0120_NO_NAME,False)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButtonIfExist("Display   \(F5\)",False)

'Capture the screenshot
Call TakeScreenShot()
Call VerifyTableCellContent(1,"Amount","SAPMV13ATCTRL_FAST_ENTRY",DT_VK11_1073_CHECK_TEXT_OF_TABLECELL_AMOUNT_0)
Call SelectRowGuiTableByRow("SAPMV13ATCTRL_FAST_ENTRY",1,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Details   \(F6\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyTextBoxContent("Lower limit","KONP-MXWRT","",DT_VK11_0300_CHECK_TEXT_OF_LOWER_LIMIT,False)
Call VerifyTextBoxContent("Upper limit","KONP-GKWRT","",DT_VK11_0300_CHECK_TEXT_OF_UPPER_LIMIT,False)

Call ClickButtonIfExist("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()
''
''''----------------------Tcode ZMDAM_BOM_REPORT----------------------------
'Enter the Tcode
Call SetTcode(DT_VK11_1073_OKCD_OCC1) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_VK11_1073_OKCD_OCC1)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Purchasing group","S_WEKGR-LOW","",DT_VK11_1000_PURCHASING_GROUP,False)
Call SetTextbox("Link Type","S_STLAN-LOW","",DT_VK11_1000_LINK_TYPE,False)
Call SetTextbox("Retention Level","S_STLAL-LOW","",DT_VK11_1000_RETENTION_LEVEL,False)
Call SetTextbox("Header Article","S_MATNR-LOW","",DT_VK11_1000_HEADER_ARTICLE,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Execute   \(F8\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call GetGridContentByTitle("","","Item Article",1,"DT_VK11_0500_GET_GRIDCELL_0_ITEM_ARTICLE_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

''
''''----------------------Tcode ZMDPC_PL_BL_PRICE_CM----------------------------
'Enter the Tcode
Call SetTcode(DT_VK11_0500_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_VK11_0500_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetComboByKey("P_ACTION",DT_VK11_1000_ACTION)
Call SetTextbox("Sales Organization","P_VKORG","",DT_VK11_1000_SALES_ORGANIZATION_OCC1,False)
Call SetTextbox("Distribution Channel","P_VTWEG","",DT_VK11_1000_DISTRIBUTION_CHANNEL_OCC1,False)
Call SetTextbox("Purchasing Group","S_EKGRP-LOW","",DT_VK11_1000_PURCHASING_GROUP_OCC1,False)
Call SetTextbox("to","S_EKGRP-HIGH","",DT_VK11_1000_TO,False)
Call SetTextbox("Article","S_MATNR-LOW","",DT_VK11_1000_ARTICLE_OCC1,False)
Call SetTextbox("Sales Price Condition Type","P_SALES","",DT_VK11_1000_SALES_PRICE_CONDITION_TYPE,False)
Call SetTextbox("Pricing Date","P_PRSDT","",DT_VK11_1000_PRICING_DATE,False)
Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Execute   \(F8\)",False)

Call SelectRowGuiGridbyRowNo("","",1,False)
'Capture the screenshot
Call TakeScreenShot()

''
''''----------------------Tcode VKP5----------------------------
'Enter the Tcode
Call SetTcode(DT_VK11_0500_OKCD_OCC1) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_VK11_0500_OKCD_OCC1)
'Capture the screenshot
Call TakeScreenShot()
'
Call ClickButtonIfExist("Get variant\.\.\.   \(Shift\+F5\)",False)

Call SetTextbox("Variant","V-LOW","",DT_VK11_0100_VARIANT,True)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call ClickButtonIfExist("Execute   \(F8\)",True)
'need to enter proper index
SAPGuiSession("transaction:=VKP5").SAPGuiWindow("transaction:=VKP5").SAPGuiButton("tooltip:=Multiple selection","index:=1").Click
Wait(2)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",1,DT_VK11_3010_TABLECELL_SINGLE_VALUE_0,True)'select
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",2,DT_VK11_3010_TABLECELL_SINGLE_VALUE_1,True)'select
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Copy   \(F8\)",True)
Wait(1)

Call SetTextbox("Validity","S_DATUM-LOW","",DT_VK11_1000_VALIDITY,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
'Capture the screenshot
Call TakeScreenShot()

''Call SelectColumnGuiGrid("","","Vendor",False)
'Call SelectColumnGuiGrid("","","Supplier",False)
'Call ClickButtonToolBar("&SORT_ASC","")

Call SetGridData("",1,"Final price",DT_VK11_0100_GRIDCELL_0_FINAL_PRICE,False)

Call GetGridContentByTitle("","","Final price",1,"DT_VK11_0100_GET_GRIDCELL_0_FINAL_PRICE_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'Capture the screenshot
Call TakeScreenShot()
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()
'save it to data sheet
Call GetStatusBar("item1","DT_VK11_1000_GET_PRICING_DOCUMENT_OUTPUT")
'reload data sheet
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'verify statusbar with datasheet feed
Call VerifyStatusBarMessageType("S")
Call VerifyStatusBar(DT_VK11_1000_CHECK_TEXT_OF_STATUSBAR)

''
''''----------------------Tcode ZMDPC_PL_BL_PRICE_CM----------------------------
'Enter the Tcode
Call SetTcode(DT_VK11_1000_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_VK11_1000_OKCD)
'Capture the screenshot
Call TakeScreenShot()


Call SetComboByKey("P_ACTION",DT_VK11_1000_ACTION_OCC1)
Call SetTextbox("Sales Organization","P_VKORG","",DT_VK11_1000_SALES_ORGANIZATION_OCC2,False)
Call SetTextbox("Distribution Channel","P_VTWEG","",DT_VK11_1000_DISTRIBUTION_CHANNEL_OCC2,False)
Call SetTextbox("Purchasing Group","S_EKGRP-LOW","",DT_VK11_1000_PURCHASING_GROUP_OCC2,False)
Call SetTextbox("to","S_EKGRP-HIGH","",DT_VK11_1000_TO_OCC1,False)
Call SetTextbox("Article","S_MATNR-LOW","",DT_VK11_1000_ARTICLE_OCC2,False)
Call SetTextbox("Sales Price Condition Type","P_SALES","",DT_VK11_1000_SALES_PRICE_CONDITION_TYPE_OCC1,False)
Call SetTextbox("Pricing Date","P_PRSDT","",DT_VK11_1000_PRICING_DATE_OCC1,False)
Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Execute   \(F8\)",False)

Call VerifyGridCellContent("",1,"PL Standard Price","",DT_VK11_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZZBRTWR_PL)

Call GetGridContentByTitle("","","Brand Leader Standard Price",1,"DT_VK13_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZZBRTWR_BL_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'Capture the screenshot
Call TakeScreenShot()


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

