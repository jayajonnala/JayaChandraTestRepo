
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRI_02_028-Check Tariff creation in Artemis_P1_VKP5_ZACTF_LEG_CODES 
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
gstrTestCaseName = "Test_S2A_PRI_02_028-Check P1_VKP5_ZACTF_LEG_CODES"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_S2A_PRI_02_028-Check Tariff creation in Artemis_P1_VKP5_ZACTF_LEG_CODES.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''''''''--------TransactionCode-VKP5----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("List Variant","P_LIVAR","",DT_VKP5_1000_LIST_VARIANT,False)
Call SetTextbox("List Group","P_LIGRU","",DT_VKP5_1000_LIST_GROUP,False)
Call SetTextbox("Article","S_MATNR-LOW","",DT_VKP5_1000_ARTICLE,False)
Call SetTextbox("Sales organization","S_VKORG-LOW","",DT_VKP5_1000_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution channel","S_VTWEG-LOW","",DT_VKP5_1000_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Price List","S_PLTYP-LOW","",DT_VKP5_1105_PRICE_LIST,False)
Call SetTextbox("Validity","S_DATUM-LOW","",ConvertDate(DT_VKP5_1000_VALIDITY),False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Enter   \(Enter\)",True)
Call TakeScreenShot
Call SelectColumnGuiGrid("", 0, "PLTYP", False)
Call ClickButtonToolBar("&MB_FILTER", 0)
Call SetTextbox("Price List","%%DYN001-LOW","",DT_VKP5_1105_PRICE_LIST,True)
Call TakeScreenShot
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot
Call SetGridData("", 1, "ENDPR", DT_VKP5_0100_GRIDCELL_0_FINAL_PRICE, False)
Call PressEnter()     
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call GetStatusBar("MessageType", "DT_VKP5_1000_GET_MESSAGE_TYPE_OUTPUT")
Call GetStatusBar("item1", "DT_VKP5_1000_GET_PRICE_CALC_NO_OUTPUT")
Call VerifyStatusBar("Data saved; pricing document "&DT_VKP5_1000_GET_PRICE_CALC_NO_OUTPUT&" created")
Call TakeScreenShot

''''''''--------TransactionCode-/nZACTF_LEG_CODES----------''''

Call SetTcode(DT_VKP5_1000_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Group Name","P_GROUP","",DT_VKP5_1000_GROUP_NAME,False)
Call SetTextbox("Article","S_MATNR-LOW","",DT_VKP5_1000_ARTICLE_OCC1,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call GetGridContent("", 0, "Legacy Value", 1, "<NA>", "<NA>", "DT_VKP5_0500_GET_GRIDCELL_0_LEGACY_VALUE_OUTPUT")
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


