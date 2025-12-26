
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_06-09-03-provide delivery note number in GR screen in Gold_check in SAP_TASE
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_06-09-03-provide Gold_check in SAP"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\DLL\RETAIL\DT_04-01-01-01-01-05-Send PO to WMS_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''--------LOGIN----------''''

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

''--------TransactionCode-ME23N----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot

Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call SetTextbox("Pur. Order","MEPO_SELECT-EBELN","",DT_ME23N_1105_CHECK_TEXT_OF_MEPO_TOPLINEEBELN,True) 
Call SelectRadioButton("MEPO_SELECT-BSTYP_F","Pur. Order",True) 
Call TakeScreenShot
Call ClickButton("Other Document   \(Enter\)",True)
Call TakeScreenShot

Call SelectTab("ITEM_DETAIL","Purchase Order History",False)
Call TakeScreenShot

Call FindRowNumber("","Short Text",DT_ME23N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BEWTK,"DT_ROWNUMBER_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_ROWNUMBER_OUTPUT",DT_ROWNUMBER)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call VerifyGridCellContent("",DT_ROWNUMBER,"Movement type",0,DT_ME23N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWART)
Call VerifyGridCellContent("",DT_ROWNUMBER,"Reference",0,DT_ME23N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_XBLNR)
Call TakeScreenShot

Call GetGridContent("",0,"Article Document",DT_ROWNUMBER,"Short Text",DT_ME23N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BEWTK,"DT_ME23N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_ME23N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_OUTPUT",DT_ME23N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call TakeScreenShot


Call LogOff()
Call FinalStatus ()



'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [12,8640008]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



