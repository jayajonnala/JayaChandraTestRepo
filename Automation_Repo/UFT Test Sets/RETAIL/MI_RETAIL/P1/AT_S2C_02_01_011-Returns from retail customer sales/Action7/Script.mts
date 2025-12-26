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

'''*****************TC1_[1]Test_S2C_02_01_011-Returns from retail customer sales_Create&Upload File_MII103_TASE***************************
TC_Path = Parameter("TC1_Path")
TC_ExcelInputFile = Parameter("TC1_ExcelinputFile")
TC_DataRowNumber1 = Parameter("TC1_1_DataRowNumber")
TC_DataRowNumber2 = Parameter("TC1_2_DataRowNumber")
TC_DataRowNumber3 = Parameter("TC1_3_DataRowNumber")
TC_DataRowNumber4 = Parameter("TC1_4_DataRowNumber")

gstrresultFolderPath= Parameter("gstrresultFolderPath")	
RunTimeResultFolder= Parameter("RunTimeResultFolder")	

Call UpdateTestPath(TC_Path)
LoadAndRunAction TC_Path,ActionName,TestIterationVal,TC_ExcelInputFile,TC_DataRowNumber1,gstrresultFolderPath,RunTimeResultFolder,TC_DataRowNumber2,TC_DataRowNumber3,TC_DataRowNumber4
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(TC_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"),RunTimeId,Priority)


