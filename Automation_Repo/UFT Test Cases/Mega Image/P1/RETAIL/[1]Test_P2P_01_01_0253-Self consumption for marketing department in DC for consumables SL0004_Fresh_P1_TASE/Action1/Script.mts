
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_0253-Self consumption for marketing department in DC for consumables SL0004_Fresh_P1     

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

gstrTestCaseName = "Test_P2P_01_01_0253_Fresh_P1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_P2P_01_01_0253-Self consumption for marketing department in DC for consumables SL0004_Fresh_P1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


''''--------TransactionCode-MB51----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call SetTextbox("Order Type","VBAK-AUART","",DT_VA01_0101_ORDER_TYPE,False)
Call SetTextbox("Sales Organization","VBAK-VKORG","",DT_VA01_0101_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution Channel","VBAK-VTWEG","",DT_VA01_0101_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Division","VBAK-SPART","",DT_VA01_0101_DIVISION,False)
Call PressEnter()
Call SetTextbox("Purch\. Order No\.","VBKD-BSTKD","",DT_VA01_4021_PO_NUMBER,False)
Call TakeScreenShot
Call SetTextbox("Sold-To Party","KUAGV-KUNNR","",DT_VA01_4701_SOLDTO_PARTY,False)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("Payment terms","VBKD-ZTERM","",DT_VA01_4440_PAYMENT_TERMS,False)
Call SetTextbox("Incoterms","VBKD-INCO1","",DT_VA01_4440_INCOTERMS,False)
Call SetComboByKey("VBAK-AUGRU", DT_VA01_4440_ORDER_REASON)
Call PressEnter()
Call SelectTab("TAXI_TABSTRIP_OVERVIEW","Item overview", False)
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG","Article",1,"","",DT_VA01_4900_TABLECELL_ARTICLE_0,False)
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG","Article",2,"","",DT_VA01_4900_TABLECELL_ARTICLE_1,False)
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG","Article",3,"","",DT_VA01_4900_TABLECELL_ARTICLE_2,False)

Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG","Order Quantity",1,"","",DT_VA01_4900_TABLECELL_ORDER_QUANTITY_0,False)
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG","Order Quantity",2,"","",DT_VA01_4900_TABLECELL_ORDER_QUANTITY_1,False)
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG","Order Quantity",3,"","",DT_VA01_4900_TABLECELL_ORDER_QUANTITY_2,False)
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Select All",false)
Call ClickButton("Item conditions",false)
Call SelectTab("TAXI_TABSTRIP_ITEM","Shipping", False)
Call TakeScreenShot
Call SetTextbox("Site","VBAP-WERKS","",DT_VA01_4452_SITE,False)
Call SetTextbox("Storage Loc\.","VBAP-LGORT","",DT_VA01_4452_STOR_LOCATION,False)
Call ClickButton("Next item   \(Shift\+F7\)",false)
Call SetTextbox("Site","VBAP-WERKS","",DT_VA01_4452_SITE_OCC1,False)
Call SetTextbox("Storage Loc\.","VBAP-LGORT","",DT_VA01_4452_STOR_LOCATION_OCC1,False)
Call ClickButton("Next item   \(Shift\+F7\)",false)
Call SetTextbox("Site","VBAP-WERKS","",DT_VA01_4452_SITE_OCC2,False)
Call SetTextbox("Storage Loc\.","VBAP-LGORT","",DT_VA01_4452_STOR_LOCATION_OCC2,False)
Call PressEnter()
Call TakeScreenShot
Call SelectTab("TAXI_TABSTRIP_ITEM","Account assignment", False)
Call TakeScreenShot
Call ClickButton("First item   \(Shift\+F5\)",false)
Call SetTextbox("Profit Center","COBL-PRCTR","",DT_VA01_1006_PROFIT_CENTER,False)
Call TakeScreenShot
Call PressEnter()
Call ClickButton("Next item   \(Shift\+F7\)",false)
Call SetTextbox("Profit Center","COBL-PRCTR","",DT_VA01_1006_PROFIT_CENTER_OCC1,False)
Call TakeScreenShot
Call PressEnter()
Call ClickButton("Next item   \(Shift\+F7\)",false)
Call SetTextbox("Profit Center","COBL-PRCTR","",DT_VA01_1006_PROFIT_CENTER_OCC2,False)
Call PressEnter()
Call ClickButton("Save   \(Ctrl\+S\)",false)
Call GetStatusBar("item2","DT_VA01_4001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call TakeScreenShot
Call VerifyStatusBar("Self Consumption "&DT_VA01_4001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" has been saved")
Call SetTcode(DT_VA01_4001_OKCD)     
Call PressEnter() 
Call SetTextbox("Order","VBAK-VBELN","",DT_VA01_4001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT,False)
Call PressEnter()
Call TakeScreenShot
Call VerifyTextBoxNoLabelContent("VBAK-VBELN",0,DT_VA01_4001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT,False)

Call LogOff()

Call FinalStatus ()














'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [8,3437477]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


