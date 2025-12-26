
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_O2C_08_07_091-Wholesale with WMS delivery - one time customer - with discount_VA01    
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


gstrTestCaseName = "Test_O2C_08_07_091-Wholesale one time customer - with discount_VA01"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_O2C_08_07_091-Wholesale with WMS delivery - one time customer - with discount_VA01.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''
'
'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

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

Call SetTextbox("Sold-To Party","KUAGV-KUNNR","",DT_VA01_4701_SOLDTO_PARTY,False)
Call PressEnter() 
Call SetTextbox("Street/House number","ADDR1_DATA-STREET","",DT_VA01_0301_STREETHOUSE_NUMBER,True)
Call SetTextbox("Street/House number","ADDR1_DATA-HOUSE_NUM1","",DT_VA01_0301_STREETHOUSE_NUMBER_OCC1,True)
Call SetTextbox("Postal Code/City","ADDR1_DATA-POST_CODE1","",DT_VA01_0301_POSTAL_CODECITY,True)
Call SetTextbox("Postal Code/City","ADDR1_DATA-CITY1","",DT_VA01_0301_POSTAL_CODECITY_OCC1,True)
Call ClickButton("G_D0100_DUMMY_NAME2",True)  ''More Fields
Call SetTextboxNoLabel("ADDR1_DATA-NAME3", "", DT_VA01_0301_ADDR1_DATANAME3, True)
Call SetTextboxNoLabel("ADDR1_DATA-NAME4", "", DT_VA01_0301_ADDR1_DATANAME4, True)
Call TakeScreenShot
Call ClickButton("Apply Entry   \(Enter\)",True)
Call SetTextbox("Delivering Site","RV45A-DWERK","",DT_VA01_4440_DELIVERING_SITE,False)
Call PressEnter()  
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Article", 1, "", "", DT_VA01_4900_TABLECELL_ARTICLE_0, false)
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Article", 2, "", "", DT_VA01_4900_TABLECELL_ARTICLE_1, false)
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Order Quantity", 1, "", "", DT_VA01_4900_TABLECELL_ORDER_QUANTITY_0, false)
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Order Quantity", 2, "", "", DT_VA01_4900_TABLECELL_ORDER_QUANTITY_1, false)
Call PressEnter() 
Call PressEnter() 
Call PressEnter() 
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG", "First date", 1, "", "", ConvertDate(DT_VL01N_4001_SELECTION_DATE_OCC1), false)
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG", "First date", 2, "", "", ConvertDate(DT_VL01N_4001_SELECTION_DATE_OCC1), false)
Call PressEnter() 
Call SelectRowGuiTableByRow("SAPMV45ATCTRL_U_ERF_AUFTRAG", 1, False)
Call ClickButtonIfExist("Item conditions",False)
Call TakeScreenShot
''''Call VerifyTableCellContent(4, "CnTy", "SAPLV69ATCTRL_KONDITIONEN", DT_VA01_6201_CHECK_TEXT_OF_TABLECELL_CNTY_3)
Call VerifyTableCellContent(5, "CnTy", "SAPLV69ATCTRL_KONDITIONEN", DT_VA01_6201_CHECK_TEXT_OF_TABLECELL_CNTY_3)
Call ClickButtonIfExist("Back   \(F3\)",False)
Call SelectRowGuiTableByRow("SAPMV45ATCTRL_U_ERF_AUFTRAG", 2, False)
Call ClickButtonIfExist("Item conditions",False)
Call TakeScreenShot
Call VerifyTableCellContent(5, "CnTy", "SAPLV69ATCTRL_KONDITIONEN", DT_VA01_6201_CHECK_TEXT_OF_TABLECELL_CNTY_3_OCC1)
Call ClickButtonIfExist("Back   \(F3\)",False)
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call GetStatusBar("item2", "DT_VA01_4001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Sales Order Push "&DT_VA01_4001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" has been saved")
Call TakeScreenShot
Call SelectMenuBar("Sales document;Deliver")
Call SetTextbox("Shipping point","LIKP-VSTEL","",DT_VL01N_4001_SHIPPING_POINT_OCC1,False)
Call SetTextbox("Selection date","LV50C-DATBI","",ConvertDate(DT_VL01N_4001_SELECTION_DATE_OCC1),False)
Call SetTextbox("Order","LV50C-VBELN","",DT_VA01_4001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT,False)
Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call GetStatusBar("item2", "DT_VL01N_4001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Outb. del.(Affiilia) "&DT_VL01N_4001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" has been saved")
Call TakeScreenShot

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



