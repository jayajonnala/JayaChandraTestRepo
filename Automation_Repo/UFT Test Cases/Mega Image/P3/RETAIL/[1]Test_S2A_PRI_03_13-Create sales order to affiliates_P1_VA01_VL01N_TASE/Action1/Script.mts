
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRI_03_13-Create sales order to affiliates_P1_VA01_VL01N
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :
'
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


gstrTestCaseName = "Test_S2A_P1_VA01_VL01N"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_S2A_PRI_03_13-Create sales order to affiliates_P1_VA01_VL01N.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'''''''''''-----Login----------'''''''
'
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''--------TransactionCode-VA01----------''''

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
Call SetTextbox("Ship-To Party","KUWEV-KUNNR","",DT_VA01_4701_SHIPTO_PARTY,False)
Call TakeScreenShot
Call SetTextbox("Delivering Site","RV45A-DWERK","",DT_VA01_4440_DELIVERING_SITE,False)
Call PressEnter() 
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Article", 1, "", "", DT_VA01_4900_TABLECELL_ARTICLE_0, false)
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Order Quantity", 1, "", "", DT_VA01_4900_TABLECELL_ORDER_QUANTITY_0, false)
Call PressEnter()
Call TakeScreenShot
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Article", 2, "", "", DT_VA01_4900_TABLECELL_ARTICLE_1, false)
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Order Quantity", 2, "", "", DT_VA01_4900_TABLECELL_ORDER_QUANTITY_1, false)
Call PressEnter()
Call TakeScreenShot
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Article", 3, "", "", DT_VA01_4900_TABLECELL_ARTICLE_2, false)
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Order Quantity", 3, "", "", DT_VA01_4900_TABLECELL_ORDER_QUANTITY_2, false)
Call PressEnter()
Call TakeScreenShot
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Article", 4, "", "", DT_VA01_4900_TABLECELL_ARTICLE_3, false)
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Order Quantity", 4, "", "", DT_VA01_4900_TABLECELL_ORDER_QUANTITY_3, false)
Call PressEnter()
Call TakeScreenShot
'Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("Select All",False)
Call ClickButtonIfExist("Item conditions",False)
Call TakeScreenShot

Call ClickButtonIfExist("Next item   \(Shift\+F7\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Next item   \(Shift\+F7\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Next item   \(Shift\+F7\)",False)
Call TakeScreenShot

Call ClickButtonIfExist("Back   \(F3\)",False)
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call GetStatusBar("item2", "DT_VA01_4001_GET_SALES_ORDER_NO_OUTPUT")
Call VerifyStatusBar("Sales Order Push "&DT_VA01_4001_GET_SALES_ORDER_NO_OUTPUT&" has been saved")
Call SelectMenuBar("Sales document;Deliver")

Call SetTextbox("Shipping point","LIKP-VSTEL","",DT_VA01_4001_SHIPPING_POINT,False)
Call SetTextbox("Selection date","LV50C-DATBI","",ConvertDate(DT_VA01_4001_SELECTION_DATE),False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call VerifyStatusBarMessageType("S")
Call GetStatusBar("item2", "DT_VA01_4001_GET_OUTBOUND_DELIVERY_NO_OUTPUT")
Call VerifyStatusBar("Outb. del.(Affiilia) "&DT_VA01_4001_GET_OUTBOUND_DELIVERY_NO_OUTPUT&" has been saved")


Call LogOff()

Call FinalStatus()






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




