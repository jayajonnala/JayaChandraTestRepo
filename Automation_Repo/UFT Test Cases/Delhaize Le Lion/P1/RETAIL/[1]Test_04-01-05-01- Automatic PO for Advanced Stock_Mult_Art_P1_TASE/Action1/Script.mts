
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_04-01-05-01- Automatic PO for Advanced Stock_Mult_Art_P1
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

If qtpParamExist("datatable_row") Then
	DataRowSet= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrTestCaseName = "Test_04-01-05-01- Automatic PO for Advanced Stock_Mult_Art_P1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\DLL\RETAIL\DT_04-03-03-01 GDSN article with Fost fresh_TASE.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario


Call LoadXMLFile(DT_LOAD_XML)

 
DT_PSTNG_DATEFormat= CSTR(Year(date)) +ConvertDoubledigit( CSTR(Month(date))) + ConvertDoubledigit(CSTR(Day(date)))
DT_DOC_DATFormat= ConvertDoubledigit(CSTR(Year(date))) +ConvertDoubledigit( Cstr(Month(date))) +  Cstr(Day(date))

Call SetNodeValue("2", "0", DT_NODE_PTD, DT_PSTNG_DATEFormat)
Call SetNodeValue("2", "0", DT_NODE_DOCD, DT_DOC_DATFormat)

Call SaveXMLFile(DT_SAVED_PATH)


Call FinalStatus ()

