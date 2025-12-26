
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_O2C_08_07_064-Order Menu maintenance - Split Order_TASE
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

gstrTestCaseName = "Test_O2C_08_07_064-Order Menu maintenance - Split Order_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_O2C_08_07_064-Order Menu maintenance - Split Order.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter() 
Call TakeScreenShot

Call SetTextbox("Sales Organization","P_VKORG","",DT_ZMDSO_ORDER_MENU_1000_SALES_ORGANIZATION,False)
Call SetTextbox("Supplying Site","S_LOCLB-LOW","",DT_ZMDSO_ORDER_MENU_1000_SUPPLYING_SITE,False)
Call SetTextbox("to","S_LOCLB-HIGH","",DT_ZMDSO_ORDER_MENU_1000_TO,False)
Call TakeScreenShot
Call SelectTab("TABSTRIP_GS_TAB", "Manual entry", False)
Call TakeScreenShot
Call SetTextbox("Distribution group","S_MVGR4-LOW","",DT_ZMDSO_ORDER_MENU_0200_DISTRIBUTION_GROUP,False)
Call SetTextbox("Valid From/on","S_VALID-LOW","",ConvertDate(DT_ZMDSO_ORDER_MENU_0200_VALID_FROMON),False)
Call SetTextbox("to","S_VALID-HIGH","",ConvertDate(DT_ZMDSO_ORDER_MENU_0200_TO),False)
Call TakeScreenShot
Call SetTextbox("Customer no. - site","S_KUNNR-LOW","",DT_ZMDSO_ORDER_MENU_0200_CUSTOMER_NO__SITE,False)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot
Call SelectColumnGuiGrid("", 0, "LOCLB", False)
Call ClickButtonToolBar("&MB_FILTER", 0)
Call SetTextbox("Sup. site","%%DYN001-LOW","",DT_FILTER_BY_SUP_SITE,True)
Call PressEnter()     
Call TakeScreenShot
Call SelectRowGuiGrid("", 0, "Sup. site", DT_FILTER_BY_SUP_SITE, False)
Call ClickButtonToolBar("ADD_BTN", 0)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot
Call SetGridData("", 1, "SDDWD1", DT_ZMDSO_ORDER_MENU_0100_GRIDCELL_0_D_THU, False)
Call TakeScreenShot
Call SetGridData("", 1, "SDPWD1", DT_ZMDSO_ORDER_MENU_0100_GRIDCELL_0__THU_OCC1, False)
Call TakeScreenShot
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot
Call SelectAllRowGuiGrid("", 0, False)
Call ClickButtonToolBar("ADD_BTN", 0)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
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




