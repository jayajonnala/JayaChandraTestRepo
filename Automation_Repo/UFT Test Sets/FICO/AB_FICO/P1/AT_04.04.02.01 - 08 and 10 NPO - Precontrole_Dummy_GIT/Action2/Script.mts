'''************ Input parameter values for Zephyr sclae result updation function****************
'These values are common for all the testsets and are mandatory to call at the start of the testset code.

TestSetName = Environment.Value("TestName")
ActionName = "AT"
TestPath = Environment.Value("TestDir")
OpCoName = StrValueOfPath(StrValueOfPathFromEnd(TestPath,2,"\"),0,"_")
TargetCycle = StrValueOfPath(StrValueOfPathFromEnd(TestPath,2,"\"),1,"_")
'Jaya Add these line of code
TestIterationVal = 0


'''*****************RunTimeResultFolderCreation_TASE***************************
RT_gstrresultFolderPath= Parameter("RT_gstrresultFolderPath")	
RT_datatable_row= Parameter("RT_datatable_row")	
RT_RunTimeResultFolder= Parameter("RT_RunTimeResultFolder")	
RT_gstrtestsetname= Parameter("RT_gstrtestsetname")	
RT_Path = Parameter("RT_Path")
gstrresultFolderPath= Parameter("gstrresultFolderPath")	
RunTimeResultFolder= Parameter("RunTimeResultFolder")	

LoadAndRunAction RT_Path,ActionName,TestIterationVal,RT_gstrresultFolderPath,RT_datatable_row,RT_RunTimeResultFolder,RT_gstrtestsetname
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(RT_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"))

