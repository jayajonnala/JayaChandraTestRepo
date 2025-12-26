	
	'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
	'.................Test Script Name :Test_06ARM02_001_Create_Returns_order_via_ARM_Franchisee_store
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
	'.................Test Script Name :Test_06ARM02_001_Create_Returns_order_via_ARM_Franchisee_store
	'.................Author : TCS 	   :Raushan
	'................ Creation Date    :20th Nov
	'.................Modified By :
	'.................Modified Date/Details :
	
	'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
	
	gstrTestCaseName = "Test_06ARM02_001_ARM_Franchisee_store"
	gstrTCDescription = ""'.....Please give TestCase Description.
	gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
	'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Direct Re- Export without border crossing_p1.xls"
	'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
	
	'Login to SAP System
	'DataRowSet=2
	Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
	Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
	Call Login(DT_SAPUSER,DT_SAPPASSWORD)
	Call PressEnter()  
	
	''----------------------Tcode VA01----------------------------
	
	Call SetTcode(DT_SAPTRANSACTIONCODE) 
	Call PressEnter()     ' 
	Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
	Call TakeScreenShot()
	
	'Enter the Outbound Delivery No
	Call SetTextbox("Order Type","VBAK-AUART","",DT_VA01_0101_ORDER_TYPE,FALSE)
	Call SetTextbox("Sales Organization","VBAK-VKORG","",DT_VA01_0101_SALES_ORGANIZATION,FALSE)
	Call SetTextbox("Distribution Channel","VBAK-VTWEG","",DT_VA01_0101_DISTRIBUTION_CHANNEL,FALSE)
	Call SetTextbox("Division","VBAK-SPART","",DT_VA01_0101_DIVISION,FALSE)
	Call TakeScreenShot()
	Call PressEnter()
	Wait(3)
	
	'Enter the details
	Call SetTextbox("Sold-To Party","KUAGV-KUNNR","",DT_VA01_4701_SOLDTO_PARTY,FALSE)
	'Call SetTextbox("Ship-To Party","KUWEV-KUNNR","",DT_VA01_4701_SHIPTO_PARTY,FALSE)
	Call TakeScreenShot()
	Call PressEnter()
	Wait(1)
	Call TakeScreenShot()
	
	Call SetTextbox("Purch\. Order No\.","VBKD-BSTKD","",DT_VA01_4021_PO_NUMBER,FALSE)
	Call SetTextbox("PO Date","VBKD-BSTDK","",ConvertDate(DT_VA01_4021_PO_DATE),FALSE)
	
	Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG","Article","1","","",DT_VA01_4900_TABLECELL_ARTICLE_0,False)
	Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG","Order Quantity","1","","",DT_VA01_4900_TABLECELL_ORDER_QUANTITY_0,False)
	Call PressEnter()
	
	'Navigate to Shipping Tab
	Call SelectTab("TAXI_TABSTRIP_OVERVIEW","Shipping",False)
	Wait(1)
	'Call PressEnter()
	Call TakeScreenShot()
	
	
	Call SetTableData("SAPMV45ATCTRL_UEIN_VERSAND","Site","1","","",DT_SITE,False)
	Call SetTableData("SAPMV45ATCTRL_UEIN_VERSAND","Shipping Point","1","","",DT_SHIPPING_POINT,False)
	Call TakeScreenShot()
	Call PressEnter()
	Wait(1)
	
	'Save the DEtails
	Call ClickButton("Save   \(Ctrl\+S\)",False) 
	Wait(2)
	Call TakeScreenShot()
	Call VerifyStatusBarMessageType("S")
	
	'Verify the Status Bar message
	Call GetStatusBar("item2","DT_ADVANCED_RETURN_ORDER_OUTPUT")
	Call VerifyStatusBarExist("AB: Advanced Returns "&DT_ADVANCED_RETURN_ORDER_OUTPUT&" has been saved")
	
	'------------------------'Log Off From Applicaton--------------------------------
	Call LogOff()
	Call FinalStatus ()
	
	'*********************************************End Of Script*********************************************************************
	

