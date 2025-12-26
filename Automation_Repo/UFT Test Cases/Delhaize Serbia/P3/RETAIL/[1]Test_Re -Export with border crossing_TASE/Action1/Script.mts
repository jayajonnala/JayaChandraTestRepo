
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Re -Export with border crossing
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
'.................Test Script Name :Test_Re -Export with border crossing
'.................Author : TCS 	   :Raushan
'................ Creation Date    :20th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Re -Export with border crossing"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Re -Export with border crossing.xls"
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
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
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
Call SetTextbox("Ship-To Party","KUWEV-KUNNR","",DT_VA01_4701_SHIPTO_PARTY,FALSE)
Call TakeScreenShot()
Call PressEnter()
Wait(1)
Call TakeScreenShot()

Call VerifyTextBoxContent("Information Message","MESSTXT1",0,lcase("Static credit check: credit limit exceeded"),True)
'Call VerifyTextBoxContent("Information Message","MESSTXT1",0,lcase("Credit check (next customer check date has passed)"),True)

'Click on Continue
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(2)

Call SetTextbox("Delivering Site","RV45A-DWERK","",DT_VA01_4440_DELIVERING_SITE,FALSE)
Call SetTextbox("Incoterms","VBKD-INCO1","",DT_VA01_4440_INCOTERMS,FALSE)
Call SetTextbox("Incoterms","VBKD-INCO2","",DT_VA01_4440_INCOTERMS_OCC1,FALSE)
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG","Article","1","","",DT_VA01_4900_TABLECELL_ARTICLE_0,False)
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG","Order Quantity","1","","",DT_VA01_4900_TABLECELL_ORDER_QUANTITY_0,False)
Call TakeScreenShot()
Call PressEnter()
Wait(1)

Call VerifyTextBoxContent("Information Message","MESSTXT1",0,lcase("Freight is redetermined"),True)

'Click on Continue
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(2)

'Navigate to the Shipping tab
Call SelectTab("TAXI_TABSTRIP_OVERVIEW","Shipping",False)
Wait(1)
Call TakeScreenShot()


'Click on Select All
Call ClickButton("Select All",False)
Call TakeScreenShot()
Wait(1)

'Click on Display item details
Call ClickButton("Display item details",False)
Call TakeScreenShot()
Wait(1)

'Navigate to the Shipping tab
Call SelectTab("TAXI_TABSTRIP_ITEM","Shipping",False)
Wait(1)

'Verify the data
Call VerifyTextBoxContent("Site","VBAP-WERKS",0,DT_VA01_4452_CHECK_TEXT_OF_SITE,False)
Call VerifyTextBoxContent("Shipping Point","VBAP-VSTEL",0,DT_VA01_4452_CHECK_TEXT_OF_SHIPPING_POINT,False)
Call VerifyTextBoxContent("Route","VBAP-ROUTE",0,DT_VA01_4452_CHECK_TEXT_OF_ROUTE,False)


'Navigate to the Conditions tab
Call SelectTab("TAXI_TABSTRIP_ITEM","Conditions",False)
Wait(1)

Call VerifyTableCellContent(2,"CnTy","SAPLV69ATCTRL_KONDITIONEN",DT_VA01_6201_CHECK_TEXT_OF_TABLECELL_CNTY_1)

'Click on Back
Call ClickButtonIfExist("Back   \(F3\)",False)
Wait(2)

'Save the DEtails
Call ClickButton("Save   \(Ctrl\+S\)",False) 
Wait(2)
Call TakeScreenShot()

Call VerifyTextBoxContent("Information Message","MESSTXT1",0,lcase("Static credit check: credit limit exceeded"),True)
'Call VerifyTextBoxContent("Information Message","MESSTXT1",0,lcase("Credit check (next customer check date has passed)"),True)

'Click on Continue
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(2)

'Click on Continue
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(2)

'Verify the Status Bar message
Call GetStatusBar("item2","DT_SALES_ORDER_PUSH")
Call VerifyStatusBar("Sales Order Push "&DT_SALES_ORDER_PUSH&" has been saved")

Call ClickButton("Exit   \(Shift\+F3\)",False)
Wait(2)

''------------------------------TCode vl01n-------------------------------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_VA01_4001_OKCD) 
Call PressEnter() 
Call CheckTCodeScreen(DT_VA01_4001_OKCD)


'Enter the Order No
Call SetTextbox("Shipping point","LIKP-VSTEL","",DT_VA01_4001_SHIPPING_POINT,False)
Call SetTextbox("Selection date","LV50C-DATBI","",ConvertDate(DT_VA01_4001_SELECTION_DATE),False)
Call SetTextbox("Order","LV50C-VBELN","",DT_VA01_4001_ORDER,False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()

'Save the DEtails
Call ClickButton("Save   \(Ctrl\+S\)",False) 
Wait(2)
Call TakeScreenShot()

'Verify the Status Bar message
Call GetStatusBar("item2","DT_OUTBOUND_ORDER")
Call VerifyStatusBar("Outb. del.(Affiilia) "&DT_OUTBOUND_ORDER&" has been saved")

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************


