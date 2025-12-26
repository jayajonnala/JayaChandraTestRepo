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

'''*****************TC1_[1]Test_S2C_05_01_049-Scrapping of products - deductible_Create&Upload File_MII103_TASE
TC_Path = Parameter("TC1_Path")
TC_ExcelInputFile = Parameter("TC1_ExcelinputFile")
TC_DataRowNumber1 = Parameter("TC1_DataRowNumber1")
TC_DataRowNumber2 = Parameter("TC1_DataRowNumber2")
TC_DataRowNumber3 = Parameter("TC1_DataRowNumber3")
TC_DataRowNumber4 = Parameter("TC1_DataRowNumber4")
TC_DataRowNumber5 = Parameter("TC1_DataRowNumber5")
TC_DataRowNumber6 = Parameter("TC1_DataRowNumber6")
TC_DataRowNumber7 = Parameter("TC1_DataRowNumber7")
TC_DataRowNumber8 = Parameter("TC1_DataRowNumber8")
TC_DataRowNumber9 = Parameter("TC1_DataRowNumber9")
TC_DataRowNumber10 = Parameter("TC1_DataRowNumber10")
TC_DataRowNumber11 = Parameter("TC1_DataRowNumber11")



gstrresultFolderPath= Parameter("gstrresultFolderPath")	
RunTimeResultFolder= Parameter("RunTimeResultFolder")	

Call UpdateTestPath(TC_Path)
LoadAndRunAction TC_Path,ActionName,TestIterationVal,TC_ExcelInputFile,TC_DataRowNumber1,gstrresultFolderPath,RunTimeResultFolder,TC_DataRowNumber2,TC_DataRowNumber3,TC_DataRowNumber4,TC_DataRowNumber5,TC_DataRowNumber6,TC_DataRowNumber7,TC_DataRowNumber8,TC_DataRowNumber9,TC_DataRowNumber10,TC_DataRowNumber11
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(TC_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"),RunTimeId,Priority)


'' Status of the Zephyr result updation 
'Call Inputparameters_ZephyrScale_Log(LogFilepath,referenceColumnName,referenceValue,targetColumnName,ExpectedValue)
'Call ValidationReporter_ZephyrScale_Log(RetriveDataFromExcelWithReference(LogFilepath,referenceColumnName,RunTimeId,targetColumnName,ExpectedValue),ExpectedValue)
