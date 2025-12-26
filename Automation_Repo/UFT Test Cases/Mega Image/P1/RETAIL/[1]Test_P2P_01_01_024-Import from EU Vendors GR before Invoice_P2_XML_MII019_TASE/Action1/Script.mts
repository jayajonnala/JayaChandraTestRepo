
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_024-Import from EU Vendors GR before Invoice_P2_XML_MII019  
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

gstrTestCaseName = "Test_P2P_01_01_024- Invoice_P2_XML_MII019"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\ssahoo\Desktop\TASEWork\Data\P1-MI-O2C\TASE_DT_O2C_08_07_092-Wholesale with WMS delivery - Ready meal customer_P2_XML_VBSS.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call LoadXMLFile(DT_XML_PATH)
Call SetNodeValue("2", "0",DT_NODE_NAME_CREATION_DATE, DT_NEW_VALUE_CREATION_DATE)
Call SetNodeValue("3", "0", DT_NODE_NAME_POSTING_DATE, DT_NEW_VALUE_POSTING_DATE)
Call SetNodeValue("3", "0",DT_NODE_NAME_DOCUMENT_DATE, DT_NEW_VALUE_DOCUMENT_DATE)
Call SetNodeValue("2", "0",DT_NODE_NAME_REF_GRP, DT_NEW_VALUE_REF_GRP)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetNodeValue("3", "0", DT_NODE_NAME_REFERENCE_NO,DT_NEW_VALUE_REFERENCE_NO)
Call SetNodeValue("3", "0", DT_NODE_NAME_MATERIAL1_NO, DT_NEW_VALUE_MATERIAL1_NO)
Call SetNodeValue("3", "1", DT_NODE_NAME_MATERIAL3_NO, DT_NEW_VALUE_MATERIAL3_NO)
Call SetNodeValue("3", "2",DT_NODE_NAME_MATERIAL4_NO, DT_NEW_VALUE_MATERIAL4_NO)
Call SetNodeValue("3", "3", DT_NODE_NAME_MATERIAL5_NO,DT_NEW_VALUE_MATERIAL5_NO)
Call SetNodeValue("3", "0", DT_NODE_NAME_PLANT1, DT_NEW_VALUE_PLANT1)
Call SetNodeValue("3", "1", DT_NODE_NAME_PLANT2, DT_NEW_VALUE_PLANT2)
Call SetNodeValue("3", "2", DT_NODE_NAME_PLANT3, DT_NEW_VALUE_PLANT3)
Call SetNodeValue("3", "3", DT_NODE_NAME_PLANT4,DT_NEW_VALUE_PLANT4)

Call SetNodeValue("3", "0", DT_NODE_NAME_VENDOR1, DT_NEW_VALUE_VENDOR1)
Call SetNodeValue("3", "1", DT_NODE_NAME_VENDOR2, DT_NEW_VALUE_VENDOR2)
Call SetNodeValue("3", "2", DT_NODE_NAME_VENDOR3, DT_NEW_VALUE_VENDOR3)
Call SetNodeValue("3", "3", DT_NODE_NAME_VENDOR4, DT_NEW_VALUE_VENDOR4)
Call SetNodeValue("3", "0", DT_NODE_NAME_QUANTITY1, DT_NEW_VALUE_QUANTITY1)
Call SetNodeValue("3", "1", DT_NODE_NAME_QUANTITY2, DT_NEW_VALUE_QUANTITY2)
Call SetNodeValue("3", "2", DT_NODE_NAME_QUANTITY3, DT_NEW_VALUE_QUANTITY3)
Call SetNodeValue("3", "3", DT_NODE_NAME_QUANTITY4, DT_NEW_VALUE_QUANTITY4)

Call SetNodeValue("3", "0", DT_NODE_NAME_PO_NUMBER1,DT_NEW_VALUE_PO_NUMBER1)
Call SetNodeValue("3", "1", DT_NODE_NAME_PO_NUMBER2,DT_NEW_VALUE_PO_NUMBER2)

Call SetNodeValue("2", "0", DT_NODE_REFINT,DT_REFINT)
Call SetNodeValue("2", "0", DT_REFMES_NAME,DT_REFMES_VALUE)
Call SaveXMLFile(DT_XML_PATH_SAVED)
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



