
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_0109-Store Non-contractual Free goods_P1_Create_Upload_MI029
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
If qtpParamExist("datatable_row1") Then
	DataRowSet1= Parameter("datatable_row1")	
End If


'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_0109-Store Non-contractual Free goods_P1_Create_Upload_MI029
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)


gstrTestCaseName = "Test_P2P_01_01_0109-Store goods_P1_Create_Upload_MI029"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_P2P_01_01_0109-Store Non-contractual Free goods_P1_Create_Upload_MI029.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_NRM",(Cint(DT_INCREMENT_NRM)+1))
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_VENDOR_DN",(Cint(DT_INCREMENT_VENDOR_DN)+1))
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_FILENAME",(Cint(DT_INCREMENT_FILENAME)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call CreateTxtFile(DT_FILE_PATH, DT_FILE_NAME, DT_END_STRING)

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet1)
Call AppendTxtFile(DT_FILE_PATH, DT_FILE_NAME, DT_END_STRING)


Call WinScpTransferFile(DT_SERVERNAME,"22",DT_USERNAME,DT_PASSWORD,DT_SOURCE_FILE,DT_DESTINATION_PATH)
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





