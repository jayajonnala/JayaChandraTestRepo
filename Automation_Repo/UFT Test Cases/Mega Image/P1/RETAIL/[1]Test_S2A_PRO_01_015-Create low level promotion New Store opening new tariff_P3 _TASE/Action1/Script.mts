
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRO_01_015-Create low level promotion New Store opening new tariff_P3
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


'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :[1]Test_P2P_01_01_0279-Duplicate invoice check_P1
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_S2A_PRO_01_015- new tariff_P3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_S2A_PRO_01_015-Create low level promotion New Store opening new tariff_P3.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)

Call Login(DT_SAPUSER,DT_SAPPASSWORD)

Call PressEnter()

Call SetTcode(DT_SAPTRANSACTIONCODE)     

Call PressEnter()     

Call TakeScreenShot()

Call SetTextBox("Condition Type","RV13A-KSCHL",0,DT_VK13_0100_CONDITION_TYPE,False)



Call ClickButton("Condition Information   \(Shift\+F4\)",False)

Call SetTextBox("Sales Organization","F002-LOW",0,DT_VK13_1000_SALES_ORGANIZATION,False)

Call SetTextBox("Distribution Channel","F003-LOW",0,DT_VK13_1000_DISTRIBUTION_CHANNEL,False)

Call SetTextBox("Article","F001-LOW",0,DT_VK13_1000_ARTICLE,False)

Call SetTextBox("from / on","SEL_DATE",0,ConvertDate(DT_VK13_1000_FROM__ON),False)

Call PressEnter()

Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)

Call TakeScreenShot()

Call SelectCheckboxNoLabel(0,"ON", False)

Call TakeScreenShot()

Call ClickButton("Display   \(F5\)",False)

Call TakeScreenShot()

Call VerifyGridCellContent("", 1, "Sales Organization", 0, DT_VK13_0010_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VKORG)

Call VerifyGridCellContent("", 1, "Distribution Channel", 0, DT_VK13_0010_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VTWEG)

Call VerifyGridCellContent("", 1, "Site", 0, DT_VK13_0010_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WERKS)

Call VerifyGridCellContent("", 1, "Article", 0,DT_VK13_0010_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR)

Call VerifyGridCellContent("", 1, "Sales Unit", 0, DT_VK13_0010_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VRKME)

Call VerifyTableCellContent(1, "CnTy", "SAPMV13ATCTRL_D0201", DT_VK13_0201_CHECK_TEXT_OF_TABLECELL_CNTY_0)

'Call VerifyTableCellContent(1, "Name", "SAPMV13ATCTRL_D0201", DT_NAME)

Call VerifyTableCellContent(1, "Amount", "SAPMV13ATCTRL_D0201", DT_VK13_0201_CHECK_TEXT_OF_TABLECELL_AMOUNT_0)

Call VerifyTableCellContent(1, "Unit", "SAPMV13ATCTRL_D0201", DT_VK13_0201_CHECK_TEXT_OF_TABLECELL_UNIT_0)

Call VerifyTableCellContent(1, "per", "SAPMV13ATCTRL_D0201", DT_VK13_0201_CHECK_TEXT_OF_TABLECELL_PER_0)

Call VerifyTableCellContent(1, "UoM", "SAPMV13ATCTRL_D0201", DT_VK13_0201_CHECK_TEXT_OF_TABLECELL_UOM_0)

Call TakeScreenShot()

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




