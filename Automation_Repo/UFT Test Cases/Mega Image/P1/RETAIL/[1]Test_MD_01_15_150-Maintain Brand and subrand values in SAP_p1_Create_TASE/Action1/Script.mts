
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_01_15_150-Maintain Brand and subrand values in SAP_p1_Create
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


gstrTestCaseName = "Test_MD_01_15_150 SAP_p1_Create"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_MD_01_15_150-Maintain Brand and subrand values in SAP.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''--------TransactionCode-/nZMDAM_BRANDS----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call ClickButton("Display \-\> Change   \(Ctrl\+F1\)",False)
Call TakeScreenShot

Call ClickButton("New Entries   \(F5\)",False)
Call TakeScreenShot

Call SetTableData("SAPLZMDAM_TAB_MAINTTCTRL_ZMDAM_V_BRAND", "Brand", 1, "", "", DT_ZMDAM_BRANDS_0200_TABLECELL_BRAND_0, False)
Call SetTableData("SAPLZMDAM_TAB_MAINTTCTRL_ZMDAM_V_BRAND", "Brand description", 1, "", "", DT_ZMDAM_BRANDS_0200_TABLECELL_BRAND_DESCRIPTION_0, False)
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot

Call SelectRowGuiTable("SAPLZMDAM_TAB_MAINTTCTRL_ZMDAM_V_BRAND", "Brand", DT_ZMDAM_BRANDS_0200_TABLECELL_BRAND_0, False)
Call ActivateItemGuiTree(0, "Brands;Sub-brands", "Sub-brands")
Call TakeScreenShot
Call ClickButton("New Entries   \(F5\)",False)
Call TakeScreenShot

Call SetTableData("SAPLZMDAM_TAB_MAINTTCTRL_ZMDAM_V_SUBBRAND", "Sub Brand", 1, "", "", DT_ZMDAM_BRANDS_0300_TABLECELL_SUB_BRAND_0, False)
Call SetTableData("SAPLZMDAM_TAB_MAINTTCTRL_ZMDAM_V_SUBBRAND", "Sub-Brand descr.", 1, "", "", DT_ZMDAM_BRANDS_0300_TABLECELL_SUBBRAND_DESCR_0, False)
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot

Call VerifyStatusBar(DT_ZMDAM_BRANDS_0300_CHECK_TEXT_OF_STATUSBAR)
Call VerifyTextBoxContent("Brand", "ZMDAM_V_SUBBRAND-BRAND", "", DT_WE09_0100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_0, False)

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



