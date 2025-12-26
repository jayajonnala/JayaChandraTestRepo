'''************ Input parameter values for Zephyr sclae result updation function****************
'These values are common for all the testsets and are mandatory to call at the start of the testset code.

TestSetName = Environment.Value("TestName")
ActionName = "AT"
TestPath = Environment.Value("TestDir")
OpCoName = StrValueOfPath(StrValueOfPath(TestPath,4,"\"),0,"_")
TargetCycle = StrValueOfPath(StrValueOfPath(TestPath,4,"\"),1,"_")
TestIterationVal = 0

'''*****************DT1_1_DataTransfer_OneValue_Optimised_TASE***************************

NumofIterations = 1 ' NUmber of iteration value to be updated according to the iterations in DT
For Iteration = 1 To NumofIterations

	DT_gstrInputExcelFilePathAndName= Parameter("DT1_"&Iteration&"_gstrInputExcelFilePathAndName")	
	DT_datatable_row= Parameter("DT1_"&Iteration&"_datatable_row")	
	DT_sourceField= Parameter("DT1_"&Iteration&"_sourceField")	
	DT_gstrTargetExcelFilePathAndName= Parameter("DT1_"&Iteration&"_gstrTargetExcelFilePathAndName")	
	DT_targetField= Parameter("DT1_"&Iteration&"_targetField")	
	DT_targetexcelrow= Parameter("DT1_"&Iteration&"_targetexcelrow")	
	DT_Path = Parameter("DT_Path")
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
	RunTimeResultFolder= Parameter("RunTimeResultFolder")
	
	LoadAndRunAction DT_Path,ActionName,TestIterationVal,DT_gstrInputExcelFilePathAndName,DT_datatable_row,DT_sourceField,DT_gstrTargetExcelFilePathAndName,DT_targetField,DT_targetexcelrow,gstrresultFolderPath,RunTimeResultFolder
	Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(DT_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"))
	
Next

