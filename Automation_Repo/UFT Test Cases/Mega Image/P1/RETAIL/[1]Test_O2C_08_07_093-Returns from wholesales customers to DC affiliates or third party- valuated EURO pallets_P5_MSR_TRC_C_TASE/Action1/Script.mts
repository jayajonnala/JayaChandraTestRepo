
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_O2C_08_07_093-Returns from wholesales customers to DC affiliates or third party- valuated EURO pallets_P5_MSR_TRC_C      
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

gstrTestCaseName = "Test_O2C_08_07_093- EURO pallets_P5_MSR_TRC_C"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_O2C_08_07_093-Returns from wholesales customers to DC affiliates or third party- valuated EURO pallets_P5_MSR_TRC_C.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


'''''--------TransactionCode-/nMSR_TRC_C----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Delivery", "SO_DLVNC-LOW", "", DT_MSR_TRC_C_0302_DELIVERY, False)
Call PressEnter() 
Call TakeScreenShot

Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "Processing Status", "", DT_MSR_TRC_C_0300_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PROC_STATUS_ICON)
Call VerifyGridCellContent("", 2, "Processing Status", "", DT_MSR_TRC_C_0300_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PROC_STATUS_ICON_OCC1)
Call ActivateNodeGuiTree(0, "#1;#7")
Call ClickButtonIfExist("Accounting overview   \(Shift\+F4\)",False)

'Call DoubleClickGuiGridCell("Documents in Accounting", 0, 1, "Doc. Number", True)
Call DoubleClickGuiGridCell("Documents in Accounting", 0, 1, "Document Number", True)

Call VerifyGridCellContent("", 1, "Amount", "", DT_MSR_TRC_C_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("", 2, "Amount", "", DT_MSR_TRC_C_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("", 3, "Amount", "", DT_MSR_TRC_C_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AZBET)
Call VerifyGridCellContent("", 1, "Currency", "", DT_MSR_TRC_C_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_RF05A_UBAZW)
Call VerifyGridCellContent("", 2, "Currency", "", DT_MSR_TRC_C_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_RF05A_UBAZW)
Call VerifyGridCellContent("", 3, "Currency", "", DT_MSR_TRC_C_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_RF05A_UBAZW)
Call VerifyGridCellContent("", 2, "Profit Center", "", DT_MSR_TRC_C_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)


Call ClickContextButtonToolBar("&MB_VARIANT", 0)
Call SelectMenuIdToolBar("&COL0",0)
Call ClickButtonToolBar("&FIND",0)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_MSR_0750_SEARCH_TERM_OCC1,True)'"Functional area"
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)
Call ClickButtonIfExist("Transfer   \(Enter\)",True)
Call TakeScreenShot


Call VerifyGridCellContent("", 2, "Functional Area", "", DT_MSR_TRC_C_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_FKBER_LONG)



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



