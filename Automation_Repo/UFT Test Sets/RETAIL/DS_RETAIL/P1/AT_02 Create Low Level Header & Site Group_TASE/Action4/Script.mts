'''************ Input parameter values for Zephyr sclae result updation function****************
'These values are common for all the testsets and are mandatory to call at the start of the testset code.

TestSetName = Environment.Value("TestName")
ActionName = "AT"
TestPath = Environment.Value("TestDir")
OpCoName = StrValueOfPath(StrValueOfPath(TestPath,4,"\"),0,"_")
TargetCycle = StrValueOfPath(StrValueOfPath(TestPath,4,"\"),1,"_")
TestIterationVal = 0

'''*****************TC1_Test_2.7.1.2.1. Create Low Level Header ZRPC_V2_TASE***************************
TC_Path = Parameter("TC1_Path")
TC_ExcelInputFile = Parameter("TC1_ExcelinputFile")
TC_DataRowNumber = Parameter("TC1_DataRowNumber")
gstrresultFolderPath= Parameter("gstrresultFolderPath")	
RunTimeResultFolder= Parameter("RunTimeResultFolder")

LoadAndRunAction TC_Path,ActionName,0,TC_ExcelInputFile,TC_DataRowNumber,gstrresultFolderPath,RunTimeResultFolder
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(TC_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"))
