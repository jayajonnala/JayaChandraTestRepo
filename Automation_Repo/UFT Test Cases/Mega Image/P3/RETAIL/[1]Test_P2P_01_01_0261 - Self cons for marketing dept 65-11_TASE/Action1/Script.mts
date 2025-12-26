
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AT_P2P_01_01_0261 - Self cons for marketing dept 65-11      
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



gstrTestCaseName = "Test_AT_P2P_01_01_0261 dept 65-11"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="TASE_DT_P2P_01_01_0261-Self consumption for marketing department in DC for consumables SL0004_P1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


''''''--------TransactionCode-VA01----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Order Type","VBAK-AUART","",DT_VA01_0101_ORDER_TYPE,False)
Call SetTextbox("Sales Organization","VBAK-VKORG","",DT_VA01_0101_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution Channel","VBAK-VTWEG","",DT_VA01_0101_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Division","VBAK-SPART","",DT_VA01_0101_DIVISION,False)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot
'DT_VA01_4021_PO_NUMBER
'Automation_Test SetTableDataNoRef(tableName, columnName, rowNumber, cellValue, blnIsItPopup)
Call SetTextbox("Purch. Order No.","VBKD-BSTKD","",DT_VA01_4021_PO_NUMBER,False)

Call SetTextbox("Sold-To Party","KUAGV-KUNNR","",DT_VA01_4701_SOLDTO_PARTY,False)
Call TakeScreenShot
Call PressEnter() 
'Call SetCombo("Order reason", "MI: Healthy eating")
Call SetComboByKey("Order reason", DT_VA01_4440_ORDER_REASON)

Call PressEnter() 

Call SetTextbox("Payment terms","VBKD-ZTERM","",DT_VA01_4440_PAYMENT_TERMS,False)
''Incoterms
Call SetTextbox("Incoterms","VBKD-INCO1","",DT_VA01_4440_INCOTERMS,False)

'Call SetTextbox("Delivering Site","RV45A-DWERK","",DT_VA01_4440_DELIVERING_SITE,False)
'Call PressEnter() 
Call SetTableDataNoRef("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Article", 1, DT_VA01_4900_TABLECELL_ARTICLE_0, false)
Call SetTableDataNoRef("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Article", 2, DT_VA01_4900_TABLECELL_ARTICLE_1, false)
Call SetTableDataNoRef("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Order Quantity", 1, DT_VA01_4900_TABLECELL_ORDER_QUANTITY_0, false)
Call SetTableDataNoRef("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Order Quantity", 2, DT_VA01_4900_TABLECELL_ORDER_QUANTITY_1, false)
Call SetTableDataNoRef("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Site", 1, DT_VA01_4900_TABLECELL_SITE_0, false)
Call SetTableDataNoRef("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Site", 2, DT_VA01_4900_TABLECELL_SITE_1, false)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

''Storage Location
Call SetTableDataNoRef("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Storage Location", 1, DT_VA01_4900_TABLECELL_STORAGE_LOCATION_0, false)
Call SetTableDataNoRef("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Storage Location", 2, DT_VA01_4900_TABLECELL_STORAGE_LOCATION_1, false)
Call PressEnter()
Call TakeScreenShot



Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call GetStatusBar("item2", "DT_VA01_4001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
'Call VerifyStatusBar("SalesOrderRetail_Aut "&DT_SO_NUMBER_OUTPUT&" has been saved")
Call VerifyStatusBar("Self Consumption "&DT_VA01_4001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" has been saved")
 

'''''''''--------TransactionCode-/nva03----------''''

Call SetTcode(DT_VA01_4001_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Order","VBAK-VBELN","",DT_VA01_4001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT,False)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot

''Validations

Call VerifyTableCellContent(1, "Article", "SAPMV45ATCTRL_U_ERF_AUFTRAG", DT_VA01_4900_CHECK_TEXT_OF_TABLECELL_ARTICLE_0)

Call VerifyTableCellContent(2, "Article", "SAPMV45ATCTRL_U_ERF_AUFTRAG", DT_VA01_4900_CHECK_TEXT_OF_TABLECELL_ARTICLE_1)

Call VerifyTextBoxContent("Sold-To Party", "KUAGV-KUNNR", "", DT_VA01_4701_CHECK_TEXT_OF_SOLDTO_PARTY, False)
	 
Call VerifyTextBoxContent("Self Consumption", "VBAK-VBELN", "", DT_VA01_4001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT, False)
	


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


