
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_O2C_05_02_050_Direct Export to foreign customer - P3_XML_Upld_VBSCN_MII81   
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

gstrTestCaseName = "Test_O2C_05_02_050_Direct Export to foreign customer - P3_XML_Upld_VBSCN_MII81"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_O2C_05_02_050_Direct Export to foreign customer - P3_Modify_Upload_VBSCN.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_DELIVERY_REFERENCE",Cint(DT_INCREMENT_DELIVERY_REFERENCE)+1)

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call LoadXMLFile(DT_DEFAULT_XML_FILE_PATH)
Call SetNodeValue("2", "0", DT_NODE_DELIVERY_NUMBER, DT_DELIVERY_NUMBER)
Call SetNodeValue("2", "0", DT_NODE_CREATION_DATE, DT_SYSTEM_DATE)
Call SetNodeValue("2", "0", DT_NODE_REFMES, DT_DELIVERY_REFERENCE_NUMBER)
Call SetNodeValue("2", "0", DT_NODE_DELIVERY_NUMBER_LINE, DT_DELIVERY_REFERENCE_NUMBER)
Call SetNodeValue("2", "0", DT_NODE_DELIV_NUMB, DT_DELIVERY_REFERENCE_NUMBER)
Call SetNodeValue("2", "1", DT_NODE_DELIV_NUMB, DT_DELIVERY_REFERENCE_NUMBER)
Call SetNodeValue("2", "2", DT_NODE_DELIV_NUMB, DT_DELIVERY_REFERENCE_NUMBER)
Call SetNodeValue("2", "3", DT_NODE_DELIV_NUMB, DT_DELIVERY_REFERENCE_NUMBER)
Call SetNodeValue("2", "4", DT_NODE_DELIV_NUMB, DT_DELIVERY_REFERENCE_NUMBER)
Call SetNodeValue("2", "5", DT_NODE_DELIV_NUMB, DT_DELIVERY_REFERENCE_NUMBER)
Call SetNodeValue("2", "6", DT_NODE_DELIV_NUMB, DT_DELIVERY_REFERENCE_NUMBER)
Call SetNodeValue("2", "7", DT_NODE_DELIV_NUMB, DT_DELIVERY_REFERENCE_NUMBER)
Call SetNodeValue("2", "0", DT_NODE_UNIT_EXID, DT_HDL_UNIT_EXID)
Call SetNodeValue("2", "0", DT_NODE_MATERIAL_1, DT_MATERIAL_NODE_VALUE1)
Call SetNodeValue("2", "1", DT_NODE_MATERIAL_2, DT_MATERIAL_NODE_VALUE2)
Call SetNodeValue("2", "0", DT_NODE_REFERENCE_NO, DT_REFINT_NUMBER)
Call SetNodeValue("2", "0", DT_NODE_TIMESTAMP, DT_TIMESTAMP_UTC)
Call SaveXMLFile(DT_SOURCE_FILE1)
Call WinScpTransferFile(DT_SERVERNAME,"22",DT_USERNAME,DT_PASSWORD,DT_SOURCE_FILE1,DT_DESTINATION_PATH)

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


