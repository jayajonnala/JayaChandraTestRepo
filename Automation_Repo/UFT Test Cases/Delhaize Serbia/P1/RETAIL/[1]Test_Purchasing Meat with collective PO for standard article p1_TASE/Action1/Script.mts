
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Purchasing Meat with collective PO for standard article p1
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

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Purchasing Meat with collective PO for standard article p1
'.................Author : TCS 	   :Raushan
'................ Creation Date    :16th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Purchasing Meat with collective PO for standard article p1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Distribution process to DS stores - SW41 (Fresh meat) - p1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode ZMDPU_COLLECT_PO----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

Call SetTextbox("Distribution center","P_WERKS","",DT_ZMDPU_COLLECT_PO_1000_DISTRIBUTION_CENTER,FALSE)
Call SetTextbox("Purch\. Organization","P_EKORG","",DT_ZMDPU_COLLECT_PO_1000_PURCH_ORGANIZATION,FALSE)
Call SetTextbox("Sales Organization","S_VKORG-LOW","",DT_ZMDPU_COLLECT_PO_1000_SALES_ORGANIZATION,FALSE)
Call SetTextbox("Distribution Channel","S_VTWEG-LOW","",DT_ZMDPU_COLLECT_PO_1000_DISTRIBUTION_CHANNEL,FALSE)
Call SetTextbox("Division","S_SPART-LOW","",DT_ZMDPU_COLLECT_PO_1000_DIVISION,FALSE)
Call SetTextbox("Creation date","S_ERDAT-LOW","",ConvertDate(DT_ZMDPU_COLLECT_PO_1000_CREATION_DATE),FALSE)
Call SetTextbox("Requested deliv\.date","S_VDATU-LOW","",ConvertDate(DT_ZMDPU_COLLECT_PO_1000_REQ_DEL_DATE),FALSE)
Call SetTextbox("Distr\. profile","S_FPRFM-LOW","",DT_ZMDPU_COLLECT_PO_1100_DISTR_PROFILE,FALSE)
Call TakeScreenShot()

Call SelectTab("TABSTRIP_TABS","CPOs creation parameters",False)
Wait(1)
Call TakeScreenShot()

'Call SelectCheckbox("P_DCCAL",0,DT_ZMDPU_COLLECT_PO_1200_DC_CALENDAR,False)

Call SetCombo("Creation Type","Collective")
Call SelectCheckbox("P_QTYUOM",0,"ON",False)
Call SetCombo("Stock Type","SAP Stock")
Call SetTextbox("Purchasing Group","P_EKGRP","",DT_ZMDPU_COLLECT_PO_1200_PURCHASING_GROUP,FALSE)
Call SetTextbox("Storage Location","P_LGORT","",DT_ZMDPU_COLLECT_PO_1200_STORAGE_LOCATION,FALSE)
Call SetTextbox("Company Code","P_BUKRS","",DT_ZMDPU_COLLECT_PO_1200_COMPANY_CODE,FALSE)
Call SetTextbox("Order Type if vendor = DC","P_BSART1","",DT_ZMDPU_COLLECT_PO_1200_ORDER_TYPE_IF_VENDOR__DC,FALSE)
Call SetTextbox("Order Type if vend=DC <> OpCo","P_BSART2","",DT_ZMDPU_COLLECT_PO_1200_ORDER_TYPE_IF_VENDDC__OPCO,FALSE)
Call SetTextbox("Order Type if vend is external","P_BSART3","",DT_ZMDPU_COLLECT_PO_1200_ORDER_TYPE_IF_VEND_IS_EXTERNAL,FALSE)
Call TakeScreenShot()

Call SelectTab("TABSTRIP_TABS","Order Parameters",False)
Wait(5)
'Enter the details
Call SelectCheckbox("P_STO",0,"ON",False)
Call TakeScreenShot()

Call SelectTab("TABSTRIP_TABS","Vendor Parameters",False)
Wait(1)
Call TakeScreenShot()
Call SetCombobykey("Vendor determination","U")
Call SetCombo("Based on","Order Delivery Date")
Call SetCombobykey("Delivery date type","V")
Call SetCombo("Delivery date check","No Check")
Call CheckScreen(DT_SAPTRANSACTIONCODE,"Collective PO Processing")
Call TakeScreenShot()

'Click on Execute
Call ClickButton("Execute   \(F8\)",False) 
Wait(2)
Call TakeScreenShot()



Call SelectRowGuiGrid("",0,"Article","7299574",False)
''Call ClickButton("Article qty per vendor   \(Shift\+F2\)",False) 
'Click on Article qty per vendor
Call ClickButton("Qty\. per vendor   \(Shift\+F2\)",False)
Wait(2)
Call TakeScreenShot()
Call ClickButton("Refresh   \(F8\)",False) 
Wait(2)
Call TakeScreenShot()



''Select grid row
'Call SelectRowGuiGrid("",0,"Vendor","10032689",False)
Call SelectRowGuiGrid("",0,"Supplier","10032689",False)
Call SetGridData("",1,"Vendor Order qty",2,False)
Call ClickButton("Refresh   \(F8\)",False) 
Wait(2)
Call TakeScreenShot()

'''Call ClickButton("Check the data   \(Ctrl\+F2\)",False)
Call ClickButton("Check   \(Ctrl\+F2\)",False)
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot()
Call PressEnter()

Call ClickButton("Save   \(Ctrl\+S\)",False) 
Wait(2)
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(3)

Call SelectRowGuiGrid("",0,"Article","7299574",False)
Call ClickButton("Create CPO   \(F7\)",False) 
Wait(2)
Call TakeScreenShot()
''Call ClickButton("Refresh   \(F8\)",False) 
''Wait(2)
''Call TakeScreenShot()

Call SelectRowGuiGrid("",0,"Article","7299574",False)
''Call ClickButton("CPOs of Vendors   \(Shift\+F4\)",False)
'Click on Article qty per vendor
Call ClickButton("Vendors CPOs   \(Shift\+F4\)",False)
Wait(2)
Call TakeScreenShot()
''Call ClickButton("Refresh   \(F8\)",False) 
''Wait(2)
''Call TakeScreenShot()

Call GetGridContent("",0,"Purchasing Document",1,"Article","7299574","DT_PURCHASING_DOC_OUTPUT")

Call ClickButton("Exit   \(Shift\+F3\)",False)
Call TakeScreenShot()
Call ClickButton("Exit   \(Shift\+F3\)",False)
Call TakeScreenShot()
Call ClickButton("Exit   \(Shift\+F3\)",False)
Call TakeScreenShot()

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************




'

