
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_2.6.2.1. Upload of purchase price condtitions via excel
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

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_2.6.2.1. Upload of purchase price condtitions via excel"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_2.6.2.1. Upload of purchase price condtitions via excel.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
'DataRowSet=6
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode ZMDPC_UPLOAD_COND ----------------------------
'Enter the transaction code
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

'Click on Get Variant Button
Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Wait(2)
Call TakeScreenShot()


Call SetTextbox("Variant","V-LOW","",DT_ZMDPC_UPLOAD_COND_0100_VARIANT,True)
Call SetTextbox("Created By","ENAME-LOW",""," ",True)
Call TakeScreenShot()

'Click on Execute
Call ClickButton("Execute   \(F8\)",True)
Wait(2)
Call TakeScreenShot()

'Enter file location
Call SetTextbox("File name","P_FILE","",DT_ZMDPC_UPLOAD_COND_1000_FILE_NAME,False)
Call TakeScreenShot()

'Click on Execute
Call ClickButton("Execute   \(F8\)",False)
Wait(2)
Call TakeScreenShot()

Call SelectAllRowGuiGrid("", 0, False)
Call TakeScreenShot()
Call ClickButton("Create Conditions   \(F8\)",False)

Call GetGridContentByTitle("", 0, "Message Text", 1, "DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MSGTX_OUTPUT")
'Call WriteRunTimeDataToExcelGlobalSheet("DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MSGTX_OUTPUT",DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MSGTX)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call GetGridContentByTitle("", 0, "Message Text", 2, "DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MSGTX_OUTPUT")
'Call WriteRunTimeDataToExcelGlobalSheet("DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MSGTX_OUTPUT",DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MSGTX)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call GetGridContentByTitle("", 0, "Message Text", 3, "DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_MSGTX_OUTPUT")
'Call WriteRunTimeDataToExcelGlobalSheet("DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_MSGTX_OUTPUT",DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MSGTX)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call GetGridContentByTitle("", 0, "Message Text", 4, "DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_MSGTX_OUTPUT")
'Call WriteRunTimeDataToExcelGlobalSheet("DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_MSGTX_OUTPUT",DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MSGTX)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call GetGridContentByTitle("", 0, "Message Text", 5, "DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_MSGTX_OUTPUT")
'Call WriteRunTimeDataToExcelGlobalSheet("DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_MSGTX_OUTPUT",DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MSGTX)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call VerifyGridCellContent("", 1, "Message Text", 0, DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MSGTX_OUTPUT)
Call VerifyGridCellContent("", 2, "Message Text", 0, DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MSGTX_OUTPUT)
Call VerifyGridCellContent("", 3, "Message Text", 0, DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_MSGTX_OUTPUT)
Call VerifyGridCellContent("", 4, "Message Text", 0, DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_MSGTX_OUTPUT)
Call VerifyGridCellContent("", 5, "Message Text", 0, DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_MSGTX_OUTPUT)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

