
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_O2C_05_02_050_Direct Export to foreign customer - P1-Create SO
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

gstrTestCaseName = "Test_O2C_05_02_050_Direct Export to foreign customer - P1-Create SO"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_O2C_05_02_050_Direct Export to foreign customer - P1-Create SO.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

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
Call SendKey("{F2}")
Call SetTextbox("Sold-To Party","KUAGV-KUNNR","",DT_VA01_4701_SOLDTO_PARTY,False)
Call TakeScreenShot
Call PressEnter() 
'Call ClickButton("Back   \(F3\)",False)
''Call ClickButtonIfExist("Back   \(F3\)",False)
Call SetTextbox("Delivering Site","RV45A-DWERK","",DT_VA01_4440_DELIVERING_SITE,False)
Call PressEnter() 
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Article", 1, "", "", DT_VA01_4900_TABLECELL_ARTICLE_0, false)
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Order Quantity", 1, "", "", DT_VA01_4900_TABLECELL_ORDER_QUANTITY_0, false)
Call PressEnter() 
Call PressEnter() 

Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Article", 2, "", "", DT_VA01_4900_TABLECELL_ARTICLE_1, false)
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG", "Order Quantity", 2, "", "", DT_VA01_4900_TABLECELL_ORDER_QUANTITY_1, false)
Call PressEnter() 
Call PressEnter()
Call ClickButtonIfExist("Select All",False)
Call SelectMenuBar("Edit;Fast change of...;Delivery date..")
Call TakeScreenShot
Call ClickButtonIfExist("Copy   \(F7\)",True)
Call TakeScreenShot
Call ClickButtonIfExist("Select All",False)
Call ClickButtonIfExist("Item conditions",False)
Call TakeScreenShot
Call ClickButtonIfExist("Next item   \(Shift\+F7\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Back   \(F3\)",False)
Call SelectMenuBar("Sales document;Deliver")
Call TakeScreenShot
'Call GetStatusBar("item2", "DT_SO_NUMBER_OUTPUT")
'Call VerifyStatusBar("Sales Order Push "&DT_SO_NUMBER_OUTPUT&" has been saved")
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetStatusBar("item2", "DT_DELIVERY_NUMBER_OUTPUT")
Call VerifyStatusBar("Outb. del.(Affiilia) "&DT_DELIVERY_NUMBER_OUTPUT&" has been saved")

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


