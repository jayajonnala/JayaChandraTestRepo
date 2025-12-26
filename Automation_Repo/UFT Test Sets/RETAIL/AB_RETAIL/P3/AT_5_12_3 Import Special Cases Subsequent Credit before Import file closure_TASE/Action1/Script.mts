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

'''*****************RunTimeResultFolderCreation_TASE***************************
RT_gstrresultFolderPath= Parameter("RT_gstrresultFolderPath")	
RT_datatable_row= Parameter("RT_datatable_row")	
RT_RunTimeResultFolder= Parameter("RT_RunTimeResultFolder")	
RT_gstrtestsetname= Parameter("RT_gstrtestsetname")	
TC_Path = Parameter("RT_Path")
gstrresultFolderPath= Parameter("gstrresultFolderPath")	
RunTimeResultFolder= Parameter("RunTimeResultFolder")	

Call UpdateTestPath(TC_Path)
LoadAndRunAction TC_Path,ActionName,TestIterationVal,RT_gstrresultFolderPath,RT_datatable_row,RT_RunTimeResultFolder,RT_gstrtestsetname
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(TC_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"), RunTimeId,Priority)

'' Status of the Zephyr result updation 
'Call Inputparameters_ZephyrScale_Log(LogFilepath,referenceColumnName,referenceValue,targetColumnName,ExpectedValue)
'Call ValidationReporter_ZephyrScale_Log(RetriveDataFromExcelWithReference(LogFilepath,referenceColumnName,RunTimeId,targetColumnName,ExpectedValue),ExpectedValue)

