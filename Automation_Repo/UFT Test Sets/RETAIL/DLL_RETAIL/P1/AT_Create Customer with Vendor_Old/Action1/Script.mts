'''************ Input parameter values for Zephyr sclae result updation function****************
'These values are common for all the testsets and are mandatory to call at the start of the testset code.

TestSetName = Environment.Value("TestName")
ActionName = Environment.Value("ActionName")
TestPath = Environment.Value("TestDir")
OpCoName = StrValueOfPath(StrValueOfPath(TestPath,4,"\"),0,"_")
TargetCycle = StrValueOfPath(StrValueOfPath(TestPath,4,"\"),1,"_")
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

'''*****************TC1_Test_03-01-01-01-01-Create Vendor_V2***************************
TC_Path = Parameter("TC1_Path")
TC_ExcelInputFile = Parameter("TC1_ExcelinputFile")
TC_DataRowNumber = Parameter("TC1_DataRowNumber")

LoadAndRunAction TC_Path,ActionName,0,TC_ExcelInputFile,TC_DataRowNumber,gstrresultFolderPath,RunTimeResultFolder
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(TC_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"))

'''*****************DT1_1_DataTransfer_OneValue_Optimised_TASE***************************
DT_gstrInputExcelFilePathAndName= Parameter("DT1_1_gstrInputExcelFilePathAndName")	
DT_datatable_row= Parameter("DT1_1_datatable_row")	
DT_sourceField= Parameter("DT1_1_sourceField")	
DT_gstrTargetExcelFilePathAndName= Parameter("DT1_1_gstrTargetExcelFilePathAndName")	
DT_targetField= Parameter("DT1_1_targetField")	
DT_targetexcelrow= Parameter("DT1_1_targetexcelrow")	
DT_Path = Parameter("DT_Path")

LoadAndRunAction DT_Path,ActionName,TestIterationVal,DT_gstrInputExcelFilePathAndName,DT_datatable_row,DT_sourceField,DT_gstrTargetExcelFilePathAndName,DT_targetField,DT_targetexcelrow,gstrresultFolderPath,RunTimeResultFolder
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(DT_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"))

'''*****************TC2_Test_03-02-01-01-02-Create Customer in Retail***************************
TC_Path = Parameter("TC2_Path")
TC_ExcelInputFile= Parameter("TC2_ExcelinputFile")	
TC_DataRowNumber= Parameter("TC2_DataRowNumber")	

LoadAndRunAction TC_Path,ActionName,0,TC_ExcelInputFile,TC_DataRowNumber,gstrresultFolderPath,RunTimeResultFolder
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(TC_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"))

'''*****************DT2_1_DataTransfer_OneValue_Optimised_TASE***************************
DT_gstrInputExcelFilePathAndName= Parameter("DT2_1_gstrInputExcelFilePathAndName")	
DT_datatable_row= Parameter("DT2_1_datatable_row")	
DT_sourceField= Parameter("DT2_1_sourceField")	
DT_gstrTargetExcelFilePathAndName= Parameter("DT2_1_gstrTargetExcelFilePathAndName")	
DT_targetField= Parameter("DT2_1_targetField")	
DT_targetexcelrow= Parameter("DT2_1_targetexcelrow")	
DT_Path = Parameter("DT_Path")

LoadAndRunAction DT_Path,ActionName,TestIterationVal,DT_gstrInputExcelFilePathAndName,DT_datatable_row,DT_sourceField,DT_gstrTargetExcelFilePathAndName,DT_targetField,DT_targetexcelrow,gstrresultFolderPath,RunTimeResultFolder
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(DT_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"))

''*****************TC3_Test_POST_03-02-01-01-02-Create Customer in Retail***************************

TC_Path  = Parameter("TC3_Path")
TC_ExcelInputFile = Parameter("TC3_ExcelinputFile")
TC_DataRowNumber = Parameter("TC3_DataRowNumber")

LoadAndRunAction TC_Path,ActionName,0,TC_ExcelInputFile,TC_DataRowNumber,gstrresultFolderPath,RunTimeResultFolder
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(TC_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"))


'''*****************DT3_1_DataTransfer_OneValue_Optimised_TASE***************************
DT_gstrInputExcelFilePathAndName= Parameter("DT3_1_gstrInputExcelFilePathAndName")	
DT_datatable_row= Parameter("DT3_1_datatable_row")	
DT_sourceField= Parameter("DT3_1_sourceField")	
DT_gstrTargetExcelFilePathAndName= Parameter("DT3_1_gstrTargetExcelFilePathAndName")	
DT_targetField= Parameter("DT3_1_targetField")	
DT_targetexcelrow= Parameter("DT3_1_targetexcelrow")	
DT_Path = Parameter("DT_Path")

LoadAndRunAction DT_Path,ActionName,TestIterationVal,DT_gstrInputExcelFilePathAndName,DT_datatable_row,DT_sourceField,DT_gstrTargetExcelFilePathAndName,DT_targetField,DT_targetexcelrow,gstrresultFolderPath,RunTimeResultFolder
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(DT_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"))

''*****************TC4_Test_POST_03-01-01-01-01-Create Vendor***************************
TC_Path  = Parameter("TC4_Path")
TC_ExcelInputFile = Parameter("TC4_ExcelinputFile")
TC_DataRowNumber = Parameter("TC4_DataRowNumber")

LoadAndRunAction TC_Path,ActionName,0,TC_ExcelInputFile,TC_DataRowNumber,gstrresultFolderPath,RunTimeResultFolder
Call PublishTestResults_PY(TestSetName,  StrValueOfLastElement(TC_Path,"\"), Environment.Value("TCExecutionStatus"), OpCoName, TargetCycle,  Environment.Value("HtmlResultReportpath"))

