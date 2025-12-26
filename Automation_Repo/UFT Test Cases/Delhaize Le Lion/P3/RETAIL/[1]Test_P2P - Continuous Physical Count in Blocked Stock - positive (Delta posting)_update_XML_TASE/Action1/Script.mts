
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_159_Update Account Completion Table (All Opcos)_TASE
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

'----------------------------------------------------------------------------------------------------------------------------
gstrTestCaseName = "Test_P2P - Continuous Physical XML"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario


Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call LoadXMLFile(DT_XML_PATH)
' SetNodeValue(nodeIndent, nodeIndex, nodeName, NewValue)
Call SetNodeValue(DT_NODE_INDENT_POSTING_DATE,DT_NODE_INDEX_POSTING_DATE,DT_NODE_NAME_POSTING_DATE,DT_NEW_VALUE_POSTING_DATE)
Call SetNodeValue(DT_NODE_INDENT_DOCUMENT_DATE,DT_NODE_INDEX_DOCUMENT_DATE,DT_NODE_NAME_DOCUMENT_DATE,DT_NEW_VALUE_DOCUMENT_DATE)

Call SaveXMLFile(DT_XML_PATH)
Call FinalStatus()

