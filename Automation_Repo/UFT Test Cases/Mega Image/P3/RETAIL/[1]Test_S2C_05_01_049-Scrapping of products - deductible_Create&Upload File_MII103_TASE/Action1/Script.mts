
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2C_05_01_049-Scrapping of products - deductible_Create&Upload File_MII103
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
''-----------------------------------------------------------------------------------------------------------------------------------------------//


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

If qtpParamExist("row3") Then
	DataRowSet2= Parameter("row3")	
End If

If qtpParamExist("row4") Then
	DataRowSet3= Parameter("row4")	
End If

If qtpParamExist("row5") Then
	DataRowSet4= Parameter("row5")	
End If
If qtpParamExist("row6") Then
	DataRowSet5= Parameter("row6")	
End If
If qtpParamExist("row7") Then
	DataRowSet6= Parameter("row7")	
End If
If qtpParamExist("row8") Then
	DataRowSet7= Parameter("row8")	
End If
If qtpParamExist("row9") Then
	DataRowSet8= Parameter("row9")	
End If
If qtpParamExist("row10") Then
	DataRowSet9= Parameter("row10")	
End If
If qtpParamExist("row11") Then
	DataRowSet10= Parameter("row11")	
End If

If qtpParamExist("RunTimeResultFolder") Then
    RunTimeResultFolder= Parameter("RunTimeResultFolder")    
End If


gstrTestCaseName = "Test_S2C_05_01_049-Scrapping of products - deductible_Create&Upload File_MII103"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_S2C_05_01_049-Scrapping of products - deductible_Create&Upload File.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

''DataRowSet = 2
''DataRowSet1 = 3
''DataRowSet2 = 4
''DataRowSet3 = 5
''DataRowSet4 = 6
''DataRowSet5 = 7
''DataRowSet6 = 8
''DataRowSet7 = 9
''DataRowSet5 = 10
''DataRowSet6 = 11
''DataRowSet7 = 12
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call CreateTxtFile(DT_FILEPATH, DT_FILENAME, DT_TEXT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet1)
Call AppendTxtFile(DT_FILEPATH, DT_FILENAME, DT_TEXT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet2)
Call AppendTxtFile(DT_FILEPATH, DT_FILENAME, DT_TEXT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet3)
Call AppendTxtFile(DT_FILEPATH, DT_FILENAME, DT_TEXT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet4)
Call AppendTxtFile(DT_FILEPATH, DT_FILENAME, DT_TEXT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet5)
Call AppendTxtFile(DT_FILEPATH, DT_FILENAME, DT_TEXT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet6)
Call AppendTxtFile(DT_FILEPATH, DT_FILENAME, DT_TEXT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet7)
Call AppendTxtFile(DT_FILEPATH, DT_FILENAME, DT_TEXT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet8)
Call AppendTxtFile(DT_FILEPATH, DT_FILENAME, DT_TEXT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet9)
Call AppendTxtFile(DT_FILEPATH, DT_FILENAME, DT_TEXT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet10)
Call AppendTxtFile(DT_FILEPATH, DT_FILENAME, DT_TEXT)



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


