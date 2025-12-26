'''************ Input parameter values for Zephyr sclae result updation function****************
'These values are common for all the testsets and are mandatory to call at the start of the testset code.

TestSetName = Environment.Value("TestName")
ActionName = "AT"
TestPath = Environment.Value("TestDir")
OpCoName = StrValueOfPath(StrValueOfPathFromEnd(TestPath,2,"\"),0,"_")
TargetCycle = StrValueOfPath(StrValueOfPathFromEnd(TestPath,2,"\"),1,"_")
Priority = StrValueOfPathFromEnd(TestPath,1,"\")
TestIterationVal = 0
RunTimeId = GenerateRunTimeID() 
Call ReporterEventRunTimeID(RunTimeId)

'''*****************DT5_1_DataTransfer_OneValue_Optimised_TASE***************************

NumofIterations = 1
For Iteration = 1 To NumofIterations

	DT_gstrInputExcelFilePathAndName= Parameter("DT5_"&Iteration&"_gstrInputExcelFilePathAndName")	
	DT_datatable_row= Parameter("DT5_"&Iteration&"_datatable_row")	
	DT_sourceField= Parameter("DT5_"&Iteration&"_sourceField")	
	DT_gstrTargetExcelFilePathAndName= Parameter("DT5_"&Iteration&"_gstrTargetExcelFilePathAndName")	
	DT_targetField= Parameter("DT5_"&Iteration&"_targetField")	
	DT_targetexcelrow= Parameter("DT5_"&Iteration&"_targetexcelrow")	
	TC_Path = Parameter("DT_Path")
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
	RunTimeResultFolder= Parameter("RunTimeResultFolder")
	
	Call UpdateTestPath(TC_Path)
	LoadAndRunAction TC_Path,ActionName,TestIterationVal,DT_gstrInputExcelFilePathAndName,DT_datatable_row,DT_sourceField,DT_gstrTargetExcelFilePathAndName,DT_targetField,DT_targetexcelrow,gstrresultFolderPath,RunTimeResultFolder
	Call CloseExcelFromTaskManager()
 	Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(TC_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"),RunTimeId,Priority)
'	 '''Status of the Zephyr result updation 
'	Call Inputparameters_ZephyrScale_Log(LogFilepath,referenceColumnName,referenceValue,targetColumnName,ExpectedValue)
'	Call ValidationReporter_ZephyrScale_Log(RetriveDataFromExcelWithReference(LogFilepath,referenceColumnName,RunTimeId,targetColumnName,ExpectedValue),ExpectedValue)
Next

'

