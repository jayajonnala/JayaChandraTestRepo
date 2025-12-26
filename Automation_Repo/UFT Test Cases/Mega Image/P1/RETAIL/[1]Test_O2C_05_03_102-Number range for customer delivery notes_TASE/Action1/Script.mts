
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_O2C_05_03_102-Number range for customer delivery notes
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


gstrTestCaseName = "Test_O2C_05_03_102-Number range for customer delivery notes"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_O2C_05_03_102-Number range for customer delivery notes.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'GetRowNo = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''-----------------Login-----------------------''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''--------TransactionCode-MB51----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Storage Location","LGORT-LOW","",DT_STORAGE_LOCATION,False)
Call SetTextbox("Movement type","BWART-LOW","",DT_MOVEMENT_TYPE,False)
Call SetTextbox("Company Code","BUKRS-LOW","",DT_COMPANY_CODE,False)
Call TakeScreenShot
Call ClickButtonIfExist("%_LGORT_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "", "", DT_STORAGE_LOCATION_2, True)
Call TakeScreenShot
Call ClickButtonIfExist("Copy   \(F8\)",False)

Call ClickButtonIfExist("%_BWART_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "", "", DT_MOVEMENT_TYPE_2, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 3, "", "", DT_MOVEMENT_TYPE_3, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 4, "", "", DT_MOVEMENT_TYPE_4, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 5, "", "", DT_MOVEMENT_TYPE_5, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 6, "", "", DT_MOVEMENT_TYPE_6, True)
Call TakeScreenShot
Call ClickButtonIfExist("Copy   \(F8\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Detail List   \(Ctrl\+Shift\+F12\)",False)
Call SelectColumnGuiGrid("", 0, "Reference", false)
Call ClickButtonIfExist("btn\[41\]",False)
Wait(5)
Call TakeScreenShot
Call SelectColumnGuiGrid("", 0, "Reference", false)
Call ClickButtonIfExist("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call TakeScreenShot
Call SendKey("{F4}")
Call TakeScreenShot
'GetLabelContentByRefLabel(refLabelContent,xDifferenceValue,yDifferenceValue,dataTableColumnName,blnIsItPopup)
Call GetLabelContentByRefLabel("Reference", 0, -32, "DT_MB51_0120_CHECK_TEXT_OF_B201800001171_OUTPUT", True)
'Call GetLabelContentByRefLabel("Reference", 0, -36, "DT_MB51_0120_CHECK_TEXT_OF_B201800001171_OUTPUT", True)

Call GetLabelContentByRefLabel("Reference", 0, -48, "DT_MB51_0120_CHECK_TEXT_OF_SECOND_OUTPUT", True)
'Call GetLabelContentByRefLabel("Reference", 0, -54, "DT_MB51_0120_CHECK_TEXT_OF_SECOND_OUTPUT", True)

Call GetLabelContentByRefLabel("Reference", 0, -64, "DT_MB51_0120_CHECK_TEXT_OF_THIRD_OUTPUT", True)
'Call GetLabelContentByRefLabel("Reference", 0, -72, "DT_MB51_0120_CHECK_TEXT_OF_THIRD_OUTPUT", True)

Call VerifyifGuiLabelExists(DT_MB51_0120_CHECK_TEXT_OF_B201800001171_OUTPUT)
Call VerifyifGuiLabelExists(DT_MB51_0120_CHECK_TEXT_OF_SECOND_OUTPUT)
Call VerifyifGuiLabelExists(DT_MB51_0120_CHECK_TEXT_OF_THIRD_OUTPUT)

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


