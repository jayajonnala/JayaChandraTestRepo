
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_SCM_Generic_Update_Upload_CAR_File_without_Archiving
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


gstrTestCaseName = "Test_SCM_Generic_Archiving"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\AB\RETAIL\DT_06DCIM03_009_Pallet_File_Quantity_Deficit_in_status_35_P01_TASE.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario



Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

 CheckFolderIndicator= checkIfFolderOrFileExists (DT_FOLDER_PATH_EXISTS,DT_FILE_NAME_EXISTS)
 Call WriteRunTimeDataToExcelGlobalSheet ("DT_FILE_FOLDER_EXISTS",CheckFolderIndicator)

If CheckFolderIndicator= TRUE Then
'	Call deleteFolder (DT_COMBINED_PATH)
	Call deleteSpecificFile(DT_PATH_TO_DELETE_FILE)
	
End If
''Call createFolder (DT_COMBINED_PATH)
Call createFolder (DT_FOLDER_PATH_EXISTS)

''Call CopyTxtFile(DT_DEFAULT_FILE_PATH,DT_DEFAULT_FILE_NAME,DT_COPY_FILE_TO_PATH)
Call CopyTxtFile(DT_DEFAULT_FILE_PATH,DT_DEFAULT_FILE_NAME,DT_FOLDER_PATH_EXISTS)
Call moveAndRenameFile(DT_FILE_PATH_AND_NAME_TO_MOVE_RENAME,DT_NEW_FILE_PATH_AND_NAME_TO_MOVE_RENAME)

Call ReplaceTxtFromFile(DT_REPLACE_TEXT_IN_FILE_PATH,DT_SEARCHED_TEXT_DATE,DT_REPLACE_TEXT_DATE)

Call ReplaceTxtFromFile(DT_REPLACE_TEXT_IN_FILE_PATH,DT_NEW_COLUMN,DT_REPLACE_TEXT_TIME)


 Call WinScpTransferFile(DT_SERVERNAME,"22",DT_USERNAME,DT_PASSWORD,DT_SOURCE_FILE,DT_DESTINATION_PATH)


Call FinalStatus ()




