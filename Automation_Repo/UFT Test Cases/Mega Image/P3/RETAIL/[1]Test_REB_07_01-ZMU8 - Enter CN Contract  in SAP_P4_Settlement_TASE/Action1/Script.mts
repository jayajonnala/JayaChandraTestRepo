
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_REB_07_01-ZMU8 - Enter CN Contract in SAP_P4_Settlement    
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




gstrTestCaseName = "Test_REB_07_01-ZMU8_P4_Settlement"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_REB_07_01-ZMU8 - Enter CN Contract  in SAP_P4_Settlement.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''-----------------------login-----------------''''

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''''''''-------Transaction Code WB2R_SV-------'''''''
'
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter() 
Call TakeScreenShot

Call SetTextbox("to","SO_DATE-HIGH","",ConvertDate(DT_WB2R_SV_1000_TO),False)
Call SetTextbox("Settlement Date","SO_DATE-LOW","",ConvertDate(DT_WB2R_SV_1000_SETTLEMENT_DATE),False)
Call SetTextbox("Posting date","P_WFDAT","",ConvertDate(DT_WB2R_SV_1000_POSTING_DATE),False)
Call SetTextbox("Document Date","P_BLDAT","",ConvertDate(DT_WB2R_SV_1000_DOCUMENT_DATE),False)
Call SetTextbox("Condition Contract","SO_NUM-LOW","",DT_WB2R_SV_1000_CONDITION_CONTRACT,False)
''Call SetTextbox("GI scrapping","P_BLDAT","",DT_MIGO_0010_GODEFAULT_TVBWART,False)
Call SetComboByKey("P_TEST", FormatBlank(DT_WB2R_SV_1000_RUN_TYPE))
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",false)
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot


Call SelectColumnGuiGrid("Condition Contract Settlement Messages", 0, "Message Text", False)
Call ClickButtonToolBar("&MB_FILTER", 0)
Call SetTextbox("Message text","%%DYN001-LOW","",DT_WB2R_SV_1105_MESSAGE_TEXT,True)
Call TakeScreenShot
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot
Call SelectRowGuiGridbyRowNo("Condition Contract Settlement Messages", 0, 1, False)

Call ClickButton("Document   \(Shift\+F9\)",False)
Call TakeScreenShot
Call ClickButton("Accounting   \(F6\)",False)
Call TakeScreenShot
'Call DoubleClickGuiGridCell("Documents in Accounting", 0, 1,"Doc. Number",True)
Call DoubleClickGuiGridCell("Documents in Accounting", 0, 1,"Document Number",True)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_WB2R_SV_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_WB2R_SV_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_WB2R_SV_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_WB2R_SV_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call TakeScreenShot

Call LogOff()

Call FinalStatus ()



'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [13,4062483]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

