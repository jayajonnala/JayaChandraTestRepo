'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_O2C_08_07_092-Wholesale with WMS delivery - Ready meal customer_P3_XML_VBSCN  
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


gstrTestCaseName = "Test_O2C_08_07_092-Wholesale  customer_P3_XML_VBSCN"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\ssahoo\Desktop\TASEWork\Data\P1-MI-O2C\TASE_DT_O2C_08_07_092-Wholesale with WMS delivery - Ready meal customer_P3_XML_VBSCN.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_FILE_NUMBER",Cint(DT_INCREMENT_FILE_NUMBER)+1)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_REFERENCE",Cint(DT_INCREMENT_REFERENCE)+1)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_REF_DOC",Cint(DT_INCREMENT_REF_DOC)+1)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_HEADER_TEXT",Cint(DT_INCREMENT_HEADER_TEXT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call LoadXMLFile(DT_DEFAULT_XML_FILE_PATH)
Call SetNodeValue("2", "0", DT_NODE_CREATION_DATE, DT_CREATION_DATE)
Call SetNodeValue("2", "0", DT_NODE_REFERENCE_NO, DT_REFERENCE_NO)
Call SetNodeValue("2", "0", DT_NODE_DELIVERY_NUMBER, DT_DELIVERY_NUMBER1)
Call SetNodeValue("2", "0", DT_NODE_REFMES, DT_REFMES)
Call SetNodeValue("2", "0", DT_NODE_DELIVERY_NUMBER_LINE, DT_DELIVERY_NUMBER_LINE)
Call SetNodeValue("2", "0", DT_NODE_DELIV_NUMB, DT_DELIV_NUMB)
Call SetNodeValue("2", "1", DT_NODE_DELIV_NUMB, DT_DELIV_NUMB)
Call SetNodeValue("2", "2", DT_NODE_DELIV_NUMB, DT_DELIV_NUMB)
Call SetNodeValue("2", "3", DT_NODE_DELIV_NUMB, DT_DELIV_NUMB)
Call SetNodeValue("2", "4", DT_NODE_DELIV_NUMB, DT_DELIV_NUMB)
Call SetNodeValue("2", "5", DT_NODE_DELIV_NUMB, DT_DELIV_NUMB)
Call SetNodeValue("2", "6", DT_NODE_DELIV_NUMB, DT_DELIV_NUMB)
Call SetNodeValue("2", "7", DT_NODE_DELIV_NUMB, DT_DELIV_NUMB)
Call SetNodeValue("2", "0", DT_NODE_UNIT_EXID, DT_UNIT_EXID)
Call SetNodeValue("2", "0", DT_NODE_TIMESTAMP, DT_TIMESTAMP)
Call SetNodeValue("2", "0", DT_NODE_MATERIAL_1, DT_MATERIAL_1)
Call SetNodeValue("2", "1", DT_NODE_MATERIAL_2, DT_MATERIAL_2)
Call SaveXMLFile(DT_TO_UPLOAD_XML_PATH)
Call WinScpTransferFile(DT_SERVERNAME,"22",DT_USERNAME,DT_PASSWORD,DT_SOURCE_FILE,DT_DESTINATION_PATH)

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



