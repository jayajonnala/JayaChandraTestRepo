
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_06-01-02-02-01-Replenish via MSI Net
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row1") Then
	DataRowSet1= Parameter("datatable_row1")	
End If
If qtpParamExist("datatable_row2") Then
	DataRowSet2= Parameter("datatable_row2")	
End If
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)


gstrTestCaseName = "Test_06-01-02-02-01-Replenish via MSI Net"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_06-01-02-02-01-Replenish via MSI Net_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//



'''Login'''
'DataRowSet1 = 2
'DataRowSet2 = 3
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet1,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

 

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet1)

Call TakeScreenshot()
Call CreateTxtFile(DT_FILEPATH, DT_FILENAME, DT_TEXT)
Call TakeScreenshot()
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet2)
Call AppendTxtFile(DT_FILEPATH, DT_FILENAME, DT_TEXT)
Call TakeScreenshot()
Call WinScpTransferFile(DT_SERVERNAME,"22",DT_USERNAME,DT_PASSWORD,DT_SOURCEFILE,DT_DESTINATION)
Call TakeScreenshot()

Call FinalStatus ()
