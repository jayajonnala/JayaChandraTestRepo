

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Direct Export to foreign customers_p1
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

gstrTestCaseName = "Test_Direct Export to foreign customers_p1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\P3\DS\RETAIL\DT_Direct Export to foreign customers_p1_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''
''----------------------VA01----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()  

Call SetTextbox("Order Type","VBAK-AUART","",DT_VA01_0101_ORDER_TYPE,False)
Call SetTextbox("Sales Organization","VBAK-VKORG","",DT_VA01_0101_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution Channel","VBAK-VTWEG","",DT_VA01_0101_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Division","VBAK-SPART","",DT_VA01_0101_DIVISION,False)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot

Call SetTextbox("Sold-To Party","KUAGV-KUNNR","",DT_VA01_4701_SOLDTO_PARTY,False)
Call SetTextbox("Ship-To Party","KUWEV-KUNNR","",DT_VA01_4701_SHIPTO_PARTY,False)
Call SetTextboxNoLabel("VBKD-INCO1","",DT_VA01_4440_INCOTERMS,False)
Call SetTextboxNoLabel("VBKD-INCO2","",DT_VA01_4440_INCOTERMS_OCC1,False)
Call SetTextbox("Delivering Site","RV45A-DWERK","",DT_VA01_4440_DELIVERING_SITE,False)
Call TakeScreenShot

Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Article", 1, "", "", DT_VA01_4900_TABLECELL_ARTICLE_0, false)
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Order Quantity", 1, "", "", DT_VA01_4900_TABLECELL_ORDER_QUANTITY_0, false)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot

Call VerifyTextBoxNoLabelContent("MESSTXT1",0,DT_VA01_0010_CHECK_TEXT_OF_MESSTXT1,True)
Call ClickButton("Continue   \(Enter\)",True)
wait 5
Call VerifyTextBoxNoLabelContent("MESSTXT1",0,DT_VA01_0010_CHECK_TEXT_OF_MESSTXT1_OCC1,True)
Call ClickButton("Continue   \(Enter\)",True)

Call ClickButton("Select All",False)
Call ClickButton("Display item details",False)
Call TakeScreenShot

Call SelectTab("TAXI_TABSTRIP_ITEM","Shipping",False)
Call TakeScreenShot
Call VerifyTextBoxContent("Site","VBAP-WERKS",0,DT_VA01_4452_CHECK_TEXT_OF_SITE,false)
Call VerifyTextBoxContent("Shipping Point","VBAP-VSTEL",0,DT_VA01_4452_CHECK_TEXT_OF_SHIPPING_POINT,false)
Call VerifyTextBoxContent("Route","VBAP-ROUTE",0,DT_VA01_4452_CHECK_TEXT_OF_ROUTE,false)

Call SelectTab("TAXI_TABSTRIP_ITEM","Conditions",False)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call VerifyTextBoxNoLabelContent("MESSTXT1",0,DT_VA01_0010_CHECK_TEXT_OF_MESSTXT1_OCC2,True)

Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot
Call GetStatusBar("item2", "DT_VA01_4001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Sales Order Push "&DT_VA01_4001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" has been saved")


'-----------------------------------VL01-----------------------------
Call SetTcode(DT_VA01_4001_OKCD) 
Call PressEnter()
Call TakeScreenShot

Call SetTextBox("Shipping point","LIKP-VSTEL",0,DT_VA01_4001_SHIPPING_POINT,False)
Call SetTextBox("Selection date","LV50C-DATBI",0,ConvertDate(DT_VA01_4001_SELECTION_DATE),False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTextBox("Order","LV50C-VBELN",0,DT_VA01_4001_ORDER,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot()
Call GetStatusBar("item2","DT_STATUSBAR_OCC1_OUTPUT")
Call VerifyStatusBar("Outb. del.(Affiilia) "&DT_STATUSBAR_OCC1_OUTPUT&" has been saved")


Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************


