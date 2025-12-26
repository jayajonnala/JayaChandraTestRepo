 
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :
'.................T-Codes Used :
'.................Author : 
'................ Creation Date : 
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Data Transfer"
gstrTCDescription = "Data Transfer"'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Delhaize_Work\DLL\P2\04-06-02-00-Goods Receipt\Input Data\DT_06-09-01-create and send manual Po in SAP.xls"
'gstrTargetExcelFilePathAndName="C:\Delhaize_Work\DLL\P2\04-06-02-00-Goods Receipt\Input Data\DT_06-09-02-Full Gr on Sloc1 (goods +pallet).xls"

If qtpParamExist("gstrInputExcelFilePathAndName") Then
gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")
End If

If qtpParamExist("gstrresultFolderPath") Then
gstrresultFolderPath= Parameter("gstrresultFolderPath")
End If

If qtpParamExist("datatable_row") Then
datatable_row= Parameter("datatable_row")
End If

If qtpParamExist("gstrTargetExcelFilePathAndName") Then
gstrTargetExcelFilePathAndName= Parameter("gstrTargetExcelFilePathAndName")
End If

If qtpParamExist("targetField") Then
targetField= Parameter("targetField")
End If

If qtpParamExist("sourceField") Then
sourceField= Parameter("sourceField")
End If


If qtpParamExist("targetexcelrow") Then
targetexcelrow= Parameter("targetexcelrow")
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
' StartExecution(excelPath, strTestCaseName, iterationIndex)


'Read the Updated Excel - Source Excel
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Output",datatable_row)

'Úpdated the Target File  - For Required Field
Call TransferDataExcel(gstrInputExcelFilePathAndName,datatable_row,sourceField,gstrTargetExcelFilePathAndName,"Global",targetField,targetexcelrow)

'Úpdated the Target File  - For Required Field
''Call TransferDataExcel (DT_ME21N_0014_CHECK_MESSAGEPARAMETER_OF_STATUSBAR,gstrTargetExcelFilePathAndName,"Global","DT_PO",3)
Call FinalStatus()
Wait 10