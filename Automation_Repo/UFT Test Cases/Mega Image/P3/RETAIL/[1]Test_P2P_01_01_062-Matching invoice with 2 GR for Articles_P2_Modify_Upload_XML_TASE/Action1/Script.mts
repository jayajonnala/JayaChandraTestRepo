'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_062-Matching invoice with 2 GR for Articles_P2_Modify_Upload_XML      
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


gstrTestCaseName = "Test_P2P_01_01_062_P2_Modify_Upload_XML"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_P2P_01_01_062-Matching invoice with 2 GR for Articles_P2_Modify_Upload_XML.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_FILE_NUMBER",Cint(DT_INCREMENT_FILE_NUMBER)+1)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_REFERENCE",Cint(DT_INCREMENT_REFERENCE)+1)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_REF_DOC",Cint(DT_INCREMENT_REF_DOC)+1)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_HEADER_TEXT",Cint(DT_INCREMENT_HEADER_TEXT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call LoadXMLFile(DT_DEFAULT_XML_FILE_PATH)
Call SetNodeValue("2", "0", DT_NODE_CREATION_DATE, DT_CREATION_DATE)
Call SetNodeValue("2", "0", DT_NODE_REFERENCE_NO, DT_REFERENCE_NO)
Call SetNodeValue("2", "0", DT_NODE_PO_NUMBER, DT_PO_NUMBER)
Call SetNodeValue("2", "0", DT_NODE_PO_NUMBER_LINE_1, DT_PO_NUMBER_LINE_1)
Call SetNodeValue("2", "1", DT_NODE_PO_NUMBER_LINE_2, DT_PO_NUMBER_LINE_2)
Call SetNodeValue("2", "0", DT_NODE_POSTING_DATE, DT_CREATION_DATE)
Call SetNodeValue("2", "0", DT_NODE_DOC_DATE, DT_CREATION_DATE)
Call SetNodeValue("2", "0", DT_NODE_REFERENCE_DOC_NO, DT_REFERENCE_DOC_NO)
Call SetNodeValue("2", "0", DT_NODE_HEADER_TEXT, DT_HEADER_TEXT)
Call SetNodeValue("2", "0", DT_NODE_MATERIAL_1, DT_MATERIAL_1)
Call SetNodeValue("2", "1", DT_NODE_MATERIAL_2, DT_MATERIAL_2)
Call SetNodeValue("2", "0", DT_NODE_PLANT_1, DT_PLANT_1)
Call SetNodeValue("2", "1", DT_NODE_PLANT_2, DT_PLANT_1)
Call SetNodeValue("2", "0", DT_NODE_VENDOR_1, DT_VENDOR_1)
Call SetNodeValue("2", "1", DT_NODE_VENDOR_2, DT_VENDOR_1)
Call SetNodeValue("2", "0", DT_NODE_QUANTITY_LINE_1, DT_QUANTITY_LINE_1)
Call SetNodeValue("2", "1", DT_NODE_QUANTITY_LINE_2, DT_QUANTITY_LINE_2)

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


