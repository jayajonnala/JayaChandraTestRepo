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

'''*****************TC3_Test_04.04.02.03 VIM - NPO Precontrole Issue - BR03 - Missing Vendor Invoice Reference (NPO)_TASE
TC_Path = Parameter("TC3_Path")
TC_ExcelInputFile = Parameter("TC3_ExcelinputFile")
TC_DataRowNumber = Parameter("TC3_DataRowNumber")
gstrresultFolderPath= Parameter("gstrresultFolderPath")	
RunTimeResultFolder= Parameter("RunTimeResultFolder")	

Call UpdateTestPath(TC_Path)
LoadAndRunAction TC_Path,ActionName,TestIterationVal,TC_ExcelInputFile,TC_DataRowNumber,gstrresultFolderPath,RunTimeResultFolder
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(TC_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"),RunTimeId,Priority)

'' Status of the Zephyr result updation 
'Call Inputparameters_ZephyrScale_Log(LogFilepath,referenceColumnName,referenceValue,targetColumnName,ExpectedValue)
'Call ValidationReporter_ZephyrScale_Log(RetriveDataFromExcelWithReference(LogFilepath,referenceColumnName,RunTimeId,targetColumnName,ExpectedValue),ExpectedValue)
