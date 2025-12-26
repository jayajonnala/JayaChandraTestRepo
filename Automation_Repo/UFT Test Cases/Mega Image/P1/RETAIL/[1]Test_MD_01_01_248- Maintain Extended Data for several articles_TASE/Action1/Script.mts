
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_01_01_248- Maintain Extended Data for several articles
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

gstrTestCaseName = "Test_MD_01_01_248- Maintain Extended Data for several articles"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\ssahoo\Desktop\TASEWork\Data\P1-MI\TASE_DT_MD_01_01_248- Maintain Extended Data for several articles.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''''--------TransactionCode-ZMDAM_MASS_MAINTAIN----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetComboByKey("Mass mainenance group", DT_ZMDAM_MASS_MAINTAIN_1000_MASS_MAINENANCE_GROUP)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
wait 2
Call ClickCellGuiGrid("Tables", "", "Choose", 3, "Table Name", "1	Nutritional Facts for Articles", False)
Call ClickCellGuiGrid("Tables", "", "Choose", 4, "Table Name", "1	Base Quantities for Nutritional facts", False)
Call ClickButton("Go to selections   \(F6\)",False)
Call TakeScreenShot
Call ClickCellTable("ZMDAM_MASS_MAINTENANCEGS_SCR0201_TBC","#7",1,"<NA>","<NA>",False)
Call TakeScreenShot
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "", "", DT_ZMDAM_MASS_MAINTAIN_3010_TABLECELL_SINGLE_VALUE_0, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "", "", DT_ZMDAM_MASS_MAINTAIN_3010_TABLECELL_SINGLE_VALUE_1, True)
Call TakeScreenShot
Call ClickButtonIfExist("Copy   \(F8\)",False)
Call SetTableData("ZMDAM_MASS_MAINTENANCEGS_SCR0201_TBC", "From Value", 2, "", "", DT_ZMDAM_MASS_MAINTAIN_0201_TABLECELL_FROM_VALUE_1, False)
Call SetTableData("ZMDAM_MASS_MAINTENANCEGS_SCR0201_TBC", "From Value", 3, "", "", DT_ZMDAM_MASS_MAINTAIN_0201_TABLECELL_FROM_VALUE_2, False)
Call TakeScreenShot
Call ClickButton("Go to maintenance   \(F6\)",False)
Call TakeScreenShot
Call SetGridData("Site Data for Article",1,"Max. Storage Period",DT_ZMDAM_MASS_MAINTAIN_0301_GRIDCELL_0_MAX_STORAGE_PERIOD,False)
Call SetGridData("Site Data for Article",2,"Max. Storage Period",DT_ZMDAM_MASS_MAINTAIN_0301_GRIDCELL_1_MAX_STORAGE_PERIOD,False)
Call SetGridData("Site Data for Article",1,"Time unit",DT_ZMDAM_MASS_MAINTAIN_0301_GRIDCELL_0_TIME_UNIT,False)
Call SetGridData("Site Data for Article",2,"Time unit",DT_ZMDAM_MASS_MAINTAIN_0301_GRIDCELL_1_TIME_UNIT,False)
Call SelectTab("GS_SCR0300_TAB", "Article extension - clie", False)
Call SetGridData("Article extension - client level",1,"Unit weight (gr)",DT_ZMDAM_MASS_MAINTAIN_0302_GRIDCELL_0_UNIT_WEIGHT__GR,False)
Call SetGridData("Article extension - client level",2,"Unit weight (gr)",DT_ZMDAM_MASS_MAINTAIN_0302_GRIDCELL_1_UNIT_WEIGHT__GR,False)
Call SetGridData("Article extension - client level",1,"Fair trade",DT_ZMDAM_MASS_MAINTAIN_0302_GRIDCELL_1_FAIR_TRADE,False)
Call SetGridData("Article extension - client level",2,"Fair trade",DT_ZMDAM_MASS_MAINTAIN_0302_GRIDCELL_1_FAIR_TRADE,False)
Call SetGridData("Article extension - client level",1,"Presence of BBD",DT_ZMDAM_MASS_MAINTAIN_0302_GRIDCELL_1_PRESENCE_OF_BBD,False)
Call SetGridData("Article extension - client level",2,"Presence of BBD",DT_ZMDAM_MASS_MAINTAIN_0302_GRIDCELL_1_PRESENCE_OF_BBD,False)
''Call SetGridData("Article extension - client level",1,"Private brand",DT_ZMDAM_MASS_MAINTAIN_0302_GRIDCELL_1_PRIVATE_BRAND,False)
''Call SetGridData("Article extension - client level",2,"Private brand",DT_ZMDAM_MASS_MAINTAIN_0302_GRIDCELL_1_PRIVATE_BRAND,False)
Call TakeScreenShot
Call PressEnter() 
Call SelectTab("GS_SCR0300_TAB", "Long Texts", False)
Call PressEnter() 
Call SetGridData("Long Texts",1,"Ingredients",DT_ZMDAM_MASS_MAINTAIN_0303_GRIDCELL_0_INGREDIENTS,False)
Call SetGridData("Long Texts",2,"Ingredients",DT_ZMDAM_MASS_MAINTAIN_0303_GRIDCELL_1_INGREDIENTS,False)
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call VerifyTextBoxNoLabelContent("DYNP_INFO","",DT_ZMDAM_MASS_MAINTAIN_0300_CHECK_TEXT_OF_DYNP_INFO,False)

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




