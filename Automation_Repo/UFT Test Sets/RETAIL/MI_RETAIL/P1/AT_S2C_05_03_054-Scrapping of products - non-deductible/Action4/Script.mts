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

'''*****************TC1_[1]Test_S2C_05_03_054-Scrapping of products - non-deductible_Create_Upld_MII103_TASE***************************
TC_Path = Parameter("TC1_Path")
TC_ExcelInputFile = Parameter("TC1_ExcelinputFile")
TC_DataRowNumber1 = Parameter("TC1_1_DataRowNumber")
TC_DataRowNumber2 = Parameter("TC1_2_DataRowNumber")
TC_DataRowNumber3 = Parameter("TC1_3_DataRowNumber")
TC_DataRowNumber4 = Parameter("TC1_4_DataRowNumber")
TC_DataRowNumber5 = Parameter("TC1_5_DataRowNumber")
TC_DataRowNumber6 = Parameter("TC1_6_DataRowNumber")
TC_DataRowNumber7 = Parameter("TC1_7_DataRowNumber")
TC_DataRowNumber8 = Parameter("TC1_8_DataRowNumber")
TC_DataRowNumber9 = Parameter("TC1_9_DataRowNumber")
TC_DataRowNumber10 = Parameter("TC1_10_DataRowNumber")
TC_DataRowNumber11 = Parameter("TC1_11_DataRowNumber")
gstrresultFolderPath= Parameter("gstrresultFolderPath")	
RunTimeResultFolder= Parameter("RunTimeResultFolder")	

Call UpdateTestPath(TC_Path)
LoadAndRunAction TC_Path,ActionName,TestIterationVal,TC_ExcelInputFile,TC_DataRowNumber1,gstrresultFolderPath,RunTimeResultFolder,TC_DataRowNumber2,TC_DataRowNumber3,TC_DataRowNumber4,TC_DataRowNumber5,TC_DataRowNumber6,TC_DataRowNumber7,TC_DataRowNumber8,TC_DataRowNumber9,TC_DataRowNumber10,TC_DataRowNumber11
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(TC_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"),RunTimeId,Priority)


