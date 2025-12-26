
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Purchasing Meat with collective PO for standard article p4
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

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_Purchasing Meat with collective PO for article with BOM - bat1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\DS\RETAIL\DT_Purchasing Meat with collective PO for article with BOM - bat1_TASE.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode MIGO----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()   

Call SetTextbox("Distribution center","P_WERKS","",DT_ZMDPU_COLLECT_PO_1000_DISTRIBUTION_CENTER,FALSE)
Call SetTextbox("Purch\. Organization","P_EKORG","",DT_ZMDPU_COLLECT_PO_1000_PURCH_ORGANIZATION,FALSE)
Call SetTextbox("Sales Organization","S_VKORG-LOW","",DT_ZMDPU_COLLECT_PO_1000_SALES_ORGANIZATION,FALSE)
Call SetTextbox("Distribution Channel","S_VTWEG-LOW","",DT_ZMDPU_COLLECT_PO_1000_DISTRIBUTION_CHANNEL,FALSE)
Call SetTextbox("Division","S_SPART-LOW","",DT_ZMDPU_COLLECT_PO_1000_DIVISION,FALSE)
Call SetTextbox("Creation date","S_ERDAT-LOW","",ConvertDate(DT_ZMDPU_COLLECT_PO_1000_CREATION_DATE1),FALSE)
Call SetTextbox("Requested deliv\.date","S_VDATU-LOW","",ConvertDate(DT_ZMDPU_COLLECT_PO_1000_CREATION_DATE),FALSE)
Call SetTextbox("Distr\. profile","S_FPRFM-LOW","",DT_ZMDPU_COLLECT_PO_1100_DISTR_PROFILE,FALSE)
Call TakeScreenShot()

Call SelectTab("TABSTRIP_TABS","CPOs creation parameters",False)
Wait(1)
Call TakeScreenShot()
Call SetTextbox("Order Type if vendor = DC","P_BSART1","",DT_ZMDPU_COLLECT_PO_1200_ORDER_TYPE_IF_VENDOR__DC,FALSE)
Call SetTextbox("Order Type if vend=DC <> OpCo","P_BSART2","",DT_ZMDPU_COLLECT_PO_1200_ORDER_TYPE_IF_VENDDC__OPCO,FALSE)
Call SetTextbox("Order Type if vend is external","P_BSART3","",DT_ZMDPU_COLLECT_PO_1200_ORDER_TYPE_IF_VEND_IS_EXTERNAL,FALSE)
Call SetTextbox("Company Code","P_BUKRS","",DT_ZMDPU_COLLECT_PO_1200_COMPANY_CODE,FALSE)
Call SetTextbox("Purchasing Group","P_EKGRP","",DT_ZMDPU_COLLECT_PO_1200_PURCHASING_GROUP,FALSE)
Call SetTextbox("Storage Location","P_LGORT","",DT_ZMDPU_COLLECT_PO_1200_STORAGE_LOCATION,FALSE)
'Call SelectCheckbox("P_DCCAL",0,"ON",False)
Call SetCombo("Creation Type","Collective")
Call SelectCheckbox("P_QTYUOM",0,"ON",False)
Call SetCombo("Stock Type","SAP Stock")
Call TakeScreenShot()
Call SelectTab("TABSTRIP_TABS","Order Parameters",False)
Wait(1)
'Enter the details
Call SelectCheckbox("P_STO",0,"ON",False)
Call TakeScreenShot()

'Navigate to the Vendor Parameters Tab
Call SelectTab("TABSTRIP_TABS","Vendor Parameters",False)
Wait(1)
Call TakeScreenShot()
Call SetCombobykey("Vendor determination","U")
Call SetCombo("Based on","Order Delivery Date")
Call SetCombobykey("Delivery date type","V")
Call SetCombo("Delivery date check","No Check")

'Select Radio button
'Call SelectRadioButton("P_USRVEN","Vendor: user's decision",False)
Call TakeScreenShot()

'Click on Execute
Call ClickButton("Execute   \(F8\)",False) 
Wait(2)
Call TakeScreenShot()
'''''''''''''''''''''''''''''''''''''change
Call SelectRowGuiGrid("",0,"Article","7299574",False)

'Click on Article qty per vendor
Call ClickButton("Qty\. per vendor   \(Shift\+F2\)",False) 
Wait(2)
Call TakeScreenShot()

'click on refresh
Call ClickButton("Refresh   \(F8\)",False) 
Wait(2)
Call TakeScreenShot()

'Call SelectRowGuiGrid("",0,"Vendor","10032689",False)
Call SelectRowGuiGrid("",0,"Supplier","10032689",False)

Call SetGridData("",1,"Vendor Order qty",2,False)

'click on refresh
Call ClickButton("Refresh   \(F8\)",False) 
Wait(2)
Call TakeScreenShot()

'Click on save
Call ClickButton("Save   \(Ctrl\+S\)",False) 
Wait(2)
Call TakeScreenShot()

Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(3)
Call VerifyGridCellContent("",1,DT_ZMDPU_COLUNM_NAME,0,DT_ZMDPU_COLLECT_PO_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LIGHT)

'Select grid row
Call SelectRowGuiGrid("",0,"Article","7299574",False)

'Click on Create CPO
Call ClickButton("Create CPO   \(F7\)",False) 
Wait(2)
Call TakeScreenShot()

'click on refresh
''Call ClickButton("Refresh   \(F8\)",False) 
Wait(2)
Call TakeScreenShot()
Call SelectRowGuiGrid("",0,"Article","7299574",False)

'Click on Article qty per vendor
Call ClickButton("Vendors CPOs   \(Shift\+F4\)",False) 
Wait(2)
Call TakeScreenShot()


''Call ClickButton("Refresh   \(F8\)",False) 
Wait(2)
Call TakeScreenShot()
Call GetGridContent("",0,"Purchasing Document",1,"Article","7299574","DT_CPO_NUMBER_OUTPUT")
Call VerifyGridCellContent("",1,"Purchasing Document",0,DT_ZMDPU_COLLECT_PO_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_CPO_NUMBER)
Call ClickCellGuiGrid("",0,"Purchasing Document",1,DT_ZMDPU_REFERENCE_COLUNM_NAME,DT_ZMDPU_REFERENCE_CELL_VALUE,False)
Call TakeScreenShot()

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

