
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2C_02_01_011-Returns from retail customer sales_P4_R1E_MB51_TASE
      
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

gstrTestCaseName = "Test_S2C_02_01_011-_P4_R1E_MB51"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_S2C_02_01_011-Returns from retail customer sales_P4_R1E_MB51.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

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
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Reference","XBLNR-LOW","",DT_MB51_1000_REFERENCE,False)
Call SetTextbox("Company Code","BUKRS-LOW","","",False)
Call SetTextbox("Site","WERKS-LOW","",DT_MB51_1000_SITE,False)
Call SetTextbox("Posting Date","BUDAT-LOW","",ConvertDate(DT_MB51_1000_POSTING_DATE),False)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call ClickButtonIfExist("Detail List   \(Ctrl\+Shift\+F12\)",False)
Call TakeScreenShot
Call DoubleClickGuiGridCell("", 0, 1, "Article Document", False)
Call SelectTab("TS_GOHEAD", "Doc. info", False)
Call ClickButtonIfExist("Close Detail Data",False)
Call TakeScreenShot
Call VerifyTableCellContent(1, "Movement type", "SAPLMIGOTV_GOITEM", DT_MB51_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_0)
Call VerifyTableCellContent(2, "Movement type", "SAPLMIGOTV_GOITEM", DT_MB51_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_1)
'Call VerifyTableCellContent(3, "Movement type", "SAPLMIGOTV_GOITEM", DT_MB51_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_2)
'Call VerifyTableCellContent(4, "Movement type", "SAPLMIGOTV_GOITEM", DT_MB51_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_3)
'Call VerifyTableCellContent(5, "Movement type", "SAPLMIGOTV_GOITEM", DT_MB51_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_4)
'Call VerifyTableCellContent(6, "Movement type", "SAPLMIGOTV_GOITEM", DT_MB51_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_5)
'Call VerifyTableCellContent(7, "Movement type", "SAPLMIGOTV_GOITEM", DT_MB51_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_6)
'Call VerifyTableCellContent(8, "Movement type", "SAPLMIGOTV_GOITEM", DT_MB51_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_7)
Call ClickButtonIfExist("Open detail data",False)
Call ClickButtonIfExist("FI Documents",False)
Call TakeScreenShot
Call SelectRowGuiGrid("Documents in Accounting", 0, "Object type text", "Accounting document", True)
Call ClickButtonIfExist("Display Document   \(F2\)",True)

Call TakeScreenShot
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_MB51_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_MB51_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 3, "BSCHL", 0, DT_MB51_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("", 4, "BSCHL", 0, DT_MB51_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BSCHL)
Call VerifyGridCellContent("", 5, "BSCHL", 0, DT_MB51_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BSCHL)
Call VerifyGridCellContent("", 6, "BSCHL", 0, DT_MB51_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_BSCHL)
Call VerifyGridCellContent("", 7, "BSCHL", 0, DT_MB51_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_6_BSCHL)
Call VerifyGridCellContent("", 8, "BSCHL", 0, DT_MB51_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_7_BSCHL)
Call VerifyGridCellContent("", 9, "BSCHL", 0, DT_MB51_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_8_BSCHL)
Call VerifyGridCellContent("", 10, "BSCHL", 0, DT_MB51_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_9_BSCHL)
Call VerifyGridCellContent("", 11, "BSCHL", 0, DT_MB51_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_10_BSCHL)
Call VerifyGridCellContent("", 12, "BSCHL", 0, DT_MB51_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_11_BSCHL)


Call VerifyGridCellContent("", 1, "KTONR", 0, DT_MB51_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_MB51_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 3, "KTONR", 0, DT_MB51_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
Call VerifyGridCellContent("", 4, "KTONR", 0, DT_MB51_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KTONR)
Call VerifyGridCellContent("", 5, "KTONR", 0, DT_MB51_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_KTONR)
Call VerifyGridCellContent("", 6, "KTONR", 0, DT_MB51_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_KTONR)
Call VerifyGridCellContent("", 7, "KTONR", 0, DT_MB51_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_6_KTONR)
Call VerifyGridCellContent("", 8, "KTONR", 0, DT_MB51_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_7_KTONR)
Call VerifyGridCellContent("", 9, "KTONR", 0, DT_MB51_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_8_KTONR)
Call VerifyGridCellContent("", 10, "KTONR", 0, DT_MB51_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_9_KTONR)
Call VerifyGridCellContent("", 11, "KTONR", 0, DT_MB51_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_10_KTONR)
Call VerifyGridCellContent("", 12, "KTONR", 0, DT_MB51_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_11_KTONR)

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


