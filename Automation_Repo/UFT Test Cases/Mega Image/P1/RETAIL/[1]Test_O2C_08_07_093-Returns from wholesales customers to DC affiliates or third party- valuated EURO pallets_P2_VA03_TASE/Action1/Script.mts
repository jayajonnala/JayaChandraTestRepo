
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_O2C_08_07_093-Returns from wholesales customers to DC affiliates or third party- valuated EURO pallets_P2_VA03     
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


gstrTestCaseName = "Test_O2C_08_07_093-Returns third party- valuated EURO pallets_P2_VA03"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_O2C_08_07_093-Returns from wholesales customers to DC affiliates or third party- valuated EURO pallets_P2_VA03.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


'''''--------TransactionCode-VA03----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Order","VBAK-VBELN","",DT_VA03_0102_ORDER,False)
Call TakeScreenShot
Call ClickButtonIfExist("Display document flow   \(Shift\+F5\)",False)
Call TakeScreenShot

Call ActivateNodeGuiTree(0,"#1;#1")
Call ClickButtonIfExist("Display document   \(F8\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Display <-> Change   \(Ctrl\+F1\)",False)
Call ClickButtonIfExist("Post Goods Receipt   \(Shift\+F8\)",False)
Call TakeScreenShot
Call GetStatusBar("item2", "DT_VA03_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Adv.Returns Delivery "&DT_VA03_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" has been saved")
Call ClickButtonIfExist("Back   \(F3\)",False)
Call ClickButtonIfExist("Back   \(F3\)",False)
Call ClickButtonIfExist("Display document flow   \(Shift\+F5\)",False)
Call TakeScreenShot
Call ActivateNodeGuiTree(0,"#1;#1;#1")
Call ClickButton("Display document   \(F8\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Back   \(F3\)",False)
Call TakeScreenShot
Call ActivateNodeGuiTree(0,"#1;#1;#1")
Call VerifyGridCellContentbyName("shell", 1, "Status", "", DT_VA03_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_STATUS)
Call GetGridContent("GD returns unrestr.*", "", "Doc.no.", 1, "<NA>", "<NA>", "DT_VA03_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DOCNUM_OUTPUT")

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



