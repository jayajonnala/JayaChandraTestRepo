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

'''*****************TC21_Test_06OM08_002_Create_Store_order_for_integrated_store_or_DC_Fresh_TASE***************************
TC_Path = Parameter("TC21_Path")
TC_ExcelInputFile = Parameter("TC21_ExcelinputFile")
TC_DataRowNumber1 = Parameter("TC21_DataRowNumber1")
TC_DataRowNumber2 = Parameter("TC21_DataRowNumber2")
TC_DataRowNumber3 = Parameter("TC21_DataRowNumber3")
TC_DataRowNumber4 = Parameter("TC21_DataRowNumber4")
TC_DataRowNumber5 = Parameter("TC21_DataRowNumber5")
TC_DataRowNumber6 = Parameter("TC21_DataRowNumber6")
TC_DataRowNumber7 = Parameter("TC21_DataRowNumber7")
TC_DataRowNumber8 = Parameter("TC21_DataRowNumber8")
TC_DataRowNumber9 = Parameter("TC21_DataRowNumber9")
TC_DataRowNumber10 = Parameter("TC21_DataRowNumber10")
TC_DataRowNumber11 = Parameter("TC21_DataRowNumber11")
TC_DataRowNumber12 = Parameter("TC21_DataRowNumber12")
TC_DataRowNumber13 = Parameter("TC21_DataRowNumber13")
TC_DataRowNumber14 = Parameter("TC21_DataRowNumber14")
TC_DataRowNumber15 = Parameter("TC21_DataRowNumber15")
TC_DataRowNumber16 = Parameter("TC21_DataRowNumber16")
TC_DataRowNumber17 = Parameter("TC21_DataRowNumber17")
TC_DataRowNumber18 = Parameter("TC21_DataRowNumber18")
TC_DataRowNumber19 = Parameter("TC21_DataRowNumber19")


gstrresultFolderPath= Parameter("gstrresultFolderPath")	
RunTimeResultFolder= Parameter("RunTimeResultFolder")	

Call UpdateTestPath(TC_Path)
LoadAndRunAction TC_Path,ActionName,TestIterationVal,TC_ExcelInputFile,TC_DataRowNumber1,gstrresultFolderPath,RunTimeResultFolder
Call CloseExcelFromTaskManager()
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(TC_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"),RunTimeId,Priority)

LoadAndRunAction TC_Path,ActionName,TestIterationVal,TC_ExcelInputFile,TC_DataRowNumber2,gstrresultFolderPath,RunTimeResultFolder
Call CloseExcelFromTaskManager()
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(TC_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"),RunTimeId,Priority)

LoadAndRunAction TC_Path,ActionName,TestIterationVal,TC_ExcelInputFile,TC_DataRowNumber3,gstrresultFolderPath,RunTimeResultFolder
Call CloseExcelFromTaskManager()
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(TC_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"),RunTimeId,Priority)

LoadAndRunAction TC_Path,ActionName,TestIterationVal,TC_ExcelInputFile,TC_DataRowNumber4,gstrresultFolderPath,RunTimeResultFolder
Call CloseExcelFromTaskManager()
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(TC_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"),RunTimeId,Priority)

LoadAndRunAction TC_Path,ActionName,TestIterationVal,TC_ExcelInputFile,TC_DataRowNumber5,gstrresultFolderPath,RunTimeResultFolder
Call CloseExcelFromTaskManager()
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(TC_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"),RunTimeId,Priority)

LoadAndRunAction TC_Path,ActionName,TestIterationVal,TC_ExcelInputFile,TC_DataRowNumber6,gstrresultFolderPath,RunTimeResultFolder
Call CloseExcelFromTaskManager()
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(TC_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"),RunTimeId,Priority)

LoadAndRunAction TC_Path,ActionName,TestIterationVal,TC_ExcelInputFile,TC_DataRowNumber7,gstrresultFolderPath,RunTimeResultFolder
Call CloseExcelFromTaskManager()
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(TC_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"),RunTimeId,Priority)

LoadAndRunAction TC_Path,ActionName,TestIterationVal,TC_ExcelInputFile,TC_DataRowNumber8,gstrresultFolderPath,RunTimeResultFolder
Call CloseExcelFromTaskManager()
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(TC_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"),RunTimeId,Priority)

LoadAndRunAction TC_Path,ActionName,TestIterationVal,TC_ExcelInputFile,TC_DataRowNumber9,gstrresultFolderPath,RunTimeResultFolder
Call CloseExcelFromTaskManager()
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(TC_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"),RunTimeId,Priority)

LoadAndRunAction TC_Path,ActionName,TestIterationVal,TC_ExcelInputFile,TC_DataRowNumber10,gstrresultFolderPath,RunTimeResultFolder
Call CloseExcelFromTaskManager()
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(TC_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"),RunTimeId,Priority)

LoadAndRunAction TC_Path,ActionName,TestIterationVal,TC_ExcelInputFile,TC_DataRowNumber11,gstrresultFolderPath,RunTimeResultFolder
Call CloseExcelFromTaskManager()
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(TC_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"),RunTimeId,Priority)

LoadAndRunAction TC_Path,ActionName,TestIterationVal,TC_ExcelInputFile,TC_DataRowNumber12,gstrresultFolderPath,RunTimeResultFolder
Call CloseExcelFromTaskManager()
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(TC_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"),RunTimeId,Priority)

LoadAndRunAction TC_Path,ActionName,TestIterationVal,TC_ExcelInputFile,TC_DataRowNumber13,gstrresultFolderPath,RunTimeResultFolder
Call CloseExcelFromTaskManager()
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(TC_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"),RunTimeId,Priority)

LoadAndRunAction TC_Path,ActionName,TestIterationVal,TC_ExcelInputFile,TC_DataRowNumber14,gstrresultFolderPath,RunTimeResultFolder
Call CloseExcelFromTaskManager()
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(TC_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"),RunTimeId,Priority)

LoadAndRunAction TC_Path,ActionName,TestIterationVal,TC_ExcelInputFile,TC_DataRowNumber15,gstrresultFolderPath,RunTimeResultFolder
Call CloseExcelFromTaskManager()
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(TC_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"),RunTimeId,Priority)

LoadAndRunAction TC_Path,ActionName,TestIterationVal,TC_ExcelInputFile,TC_DataRowNumber16,gstrresultFolderPath,RunTimeResultFolder
Call CloseExcelFromTaskManager()
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(TC_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"),RunTimeId,Priority)

LoadAndRunAction TC_Path,ActionName,TestIterationVal,TC_ExcelInputFile,TC_DataRowNumber17,gstrresultFolderPath,RunTimeResultFolder
Call CloseExcelFromTaskManager()
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(TC_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"),RunTimeId,Priority)

LoadAndRunAction TC_Path,ActionName,TestIterationVal,TC_ExcelInputFile,TC_DataRowNumber18,gstrresultFolderPath,RunTimeResultFolder
Call CloseExcelFromTaskManager()
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(TC_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"),RunTimeId,Priority)

LoadAndRunAction TC_Path,ActionName,TestIterationVal,TC_ExcelInputFile,TC_DataRowNumber19,gstrresultFolderPath,RunTimeResultFolder
Call CloseExcelFromTaskManager()
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(TC_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"),RunTimeId,Priority)

'' Status of the Zephyr result updation 
'Call Inputparameters_ZephyrScale_Log(LogFilepath,referenceColumnName,referenceValue,targetColumnName,ExpectedValue)
'Call ValidationReporter_ZephyrScale_Log(RetriveDataFromExcelWithReference(LogFilepath,referenceColumnName,RunTimeId,targetColumnName,ExpectedValue),ExpectedValue)
