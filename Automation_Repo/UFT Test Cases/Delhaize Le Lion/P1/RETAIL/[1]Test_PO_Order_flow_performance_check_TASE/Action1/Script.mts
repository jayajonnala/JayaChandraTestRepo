

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_PO Order flow_Winscp_Transfer_TASE
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
	DataRowSet1= Parameter("datatable_row_1")
End If

If qtpParamExist("RunTimeResultFolder") Then
    RunTimeResultFolder= Parameter("RunTimeResultFolder")    
End If


gstrTestCaseName = "Test_PO_Order_flow_performance_check_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet1,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''
''Reading MSI existing value from first branch file
Call ReadTextByLineNumber(DT_SOURCE_FILE_1,DT_LINE_NUMBER,"DT_MSI_VALUE_FROM_FILE_OUTPUT")
''Reading ICAROS existing value from third branch file
Call ReadTextByLineNumber(DT_SOURCE_FILE_3,DT_LINE_NUMBER,"DT_ICAROS_VALUE_FROM_FILE_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet1)
''Incrementing the MSI value obtained from text file 
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_1",(Cint(DT_EXISTING_MSI_VALUE)+1))
''Incrementing the ICAROS value obtained from text file 
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_2",(Cint(DT_EXISTING_ICAROS_VALUE)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet1)
''In the first branch file, replacing the existing MSI value with incremented value
Call ReplaceTxtFromFile(DT_SOURCE_FILE_1,DT_EXISTING_TEXT_MSI,DT_TEXT_TO_REPLACE_1)
''In the second branch file, replacing the existing MSI value with incremented value
''In the second branch file, replacing the existing ICAROS value with incremented value
Call ReplaceTxtFromFile(DT_SOURCE_FILE_2,DT_EXISTING_TEXT_MSI,DT_TEXT_TO_REPLACE_1)
Call ReplaceTxtFromFile(DT_SOURCE_FILE_2,DT_EXISTING_TEXT_ICAROS,DT_TEXT_TO_REPLACE_2)
''In the third branch file, replacing the existing ICAROS value with incremented value
Call ReplaceTxtFromFile(DT_SOURCE_FILE_3,DT_EXISTING_TEXT_ICAROS,DT_TEXT_TO_REPLACE_2)
''Capturing the Winscp transfer commencing time
Call WriteRunTimeDataToExcel("DT_WINSCP_TRANSFER_STARTING_TIME_OUTPUT", time)
'''Winscp transfer is performed for first branch file
Call WinScpTransferFile(DT_SERVER_NAME,"22",DT_USER_NAME,DT_PASSWORD,DT_SOURCE_FILE_1,DT_TARGET)
''Winscp transfer is performed for second branch file
Call WinScpTransferFile(DT_SERVER_NAME,"22",DT_USER_NAME,DT_PASSWORD,DT_SOURCE_FILE_2,DT_TARGET)
''Winscp transfer is performed for third branch file
Call WinScpTransferFile(DT_SERVER_NAME,"22",DT_USER_NAME,DT_PASSWORD,DT_SOURCE_FILE_3,DT_TARGET)
''Capturing the Winscp transfer completion time
Call WriteRunTimeDataToExcel("DT_WINSCP_TRANSFER_COMPLETION_TIME_OUTPUT", time)

'' After Winscp transfer, waiting for some time (Execution threshold) before checking idocs in SAP
wait(CInt(DT_EXECUTION_WAIT_TIME))

''WE02 - verify idocs in SAP
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet1)

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
AIUtil.SetContext SAPGuisession(sessionObject)
'
'''--------TransactionCode-WE02----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SelectTab("TABSTRIP_IDOCTABBL", "EDI", False)
Call ClickButtonByIndex("%_REFINT_%_APP_%-VALU_PUSH", 1, False)
Call SetTableDataNoRef("SAPLALDBSINGLE", "Single value", 1, DT_WE02_1300_TRANSFER_FILE_REFERENCE_1, True)
Call SetTableDataNoRef("SAPLALDBSINGLE", "Single value", 2, DT_WE02_1300_TRANSFER_FILE_REFERENCE_2, True)
Call TakeScreenShot
Call ClickButton("Copy   \(F8\)",True)
Call ClickButtonByIndex("%_REFMES_%_APP_%-VALU_PUSH", 1, False)
Call SetTableDataNoRef("SAPLALDBSINGLE", "Single value", 1, DT_WE02_1300_MESSAGE_REFERENCE_WITH_PREFIX, True)
Call SetTableDataNoRef("SAPLALDBSINGLE", "Single value", 2, DT_WE02_1300_MESSAGE_REFERENCE_2_WITH_SUFFIX, True)
Call TakeScreenShot
Call ClickButton("Copy   \(F8\)",True)
Call TakeScreenShot
Call SelectTab("TABSTRIP_IDOCTABBL", "Default", False)
Call SetTextbox("Created On","CREDAT-LOW","",ConvertDate(DT_WE02_1100_CREATED_ON),False)
Call SetTextbox("to","CREDAT-HIGH","",ConvertDate(DT_WE02_1100_TO_OCC1),False)
Call SetTextbox("Created At","CRETIM-LOW","",DT_WE02_1100_CREATED_AT_CET,False)
'Call SetTextbox("to","CRETIM-HIGH","",DT_WE02_1100_TO_CET,False)
Call SetTextbox("Message Variant","MESCOD-LOW","",DT_WE02_1100_MESSAGE_VARIANT,False)
Call SetTextbox("Message Function","MESFCT-LOW","",DT_WE02_1100_MESSAGE_FUNCTION,False)
Call TakeScreenShot
DT_WE02_1100_TO_CET=CDate(DT_WINSCP_TRANSFER_COMPLETION_TIME_CET) 
existingVal=0
For Iterator = 1 To CInt(DT_EXECUTION_THRESHOLD+1) Step 1
	
	DT_WE02_1100_TO_CET=CDate(CDate(DT_WE02_1100_TO_CET) + TimeValue("00:01:00"))
	Call SetTextbox("to","CRETIM-HIGH","",DT_WE02_1100_TO_CET,False)
	Call TakeScreenShot
	Call ClickButton("Execute   \(F8\)",False)
	'Call TakeScreenShot
	If SAPGuiSession(sessionObject).SapGuiWindow(windowobject).SAPGuiButton("guicomponenttype:=40","tooltip:=Continue   \(Enter\)").Exist(5)=True Then
		Call ClickButton("Continue   \(Enter\)",True)
		Call TakeScreenShot()
	Else
		'Call TakeScreenShot()
		AIUtil("right_triangle", micAnyText, micFromTop, 5).Click
		Call TakeScreenShot()
		time_diff=CInt(DateDiff("n",DT_WINSCP_TRANSFER_COMPLETION_TIME_CET,DT_WE02_1100_TO_CET))
		Call GetNumberOfRows("shell", 0, "DT_NO_OF_IDOCS_GENERATED_OUTPUT")
		Call TakeScreenShot()
		msg="In "&time_diff&" minutes, '"&DT_NO_OF_IDOCS_GENERATED_OUTPUT&" idocs' generated"
		Call UpdateResultToReport("Idoc generation status","",msg,"DONE","")
	End  If
	If existingVal=DT_NO_OF_IDOCS_GENERATED_OUTPUT and DT_NO_OF_IDOCS_GENERATED_OUTPUT>=Cint(DT_THRESHOLD_VALUE) Then
		Exit For
	End If	
	existingVal=DT_NO_OF_IDOCS_GENERATED_OUTPUT
	Call ClickButton("Back   \(F3\)",False)
Next

total_time_taken=CInt(DateDiff("n",DT_WINSCP_TRANSFER_COMPLETION_TIME_CET,DT_WE02_1100_TO_CET))-1
threshold_time=CInt(Mid(CStr(DT_THRESHOLD_WAIT),4,2))

If total_time_taken<=threshold_time and DT_NO_OF_IDOCS_GENERATED_OUTPUT>=Cint(DT_THRESHOLD_VALUE) Then
	strMsg="PO Connection is healthy and '"&DT_NO_OF_IDOCS_GENERATED_OUTPUT&" idocs' generated within '"&total_time_taken&" minutes'(threshold time)"
	strStatus = "PASS"
ElseIf total_time_taken>threshold_time and DT_NO_OF_IDOCS_GENERATED_OUTPUT>=Cint(DT_THRESHOLD_VALUE) Then
 	strMsg="Idocs not generated within "&threshold_time&" minutes and took "&total_time_taken&" minutes to generate "&DT_NO_OF_IDOCS_GENERATED_OUTPUT&" idocs"
 	strStatus="FAIL"

ElseIf DT_NO_OF_IDOCS_GENERATED_OUTPUT<Cint(DT_THRESHOLD_VALUE) Then
       strMsg="Expected number of "&DT_THRESHOLD_VALUE& " Idocs not generated even after "&total_time_taken&" minutes"
 	strStatus="FAIL"
End If

Call UpdateResultToReport("PO Connection Health Status","",strMsg,strStatus,"")

Set existingVal=Nothing
Set DT_WE02_1100_TO_CET=Nothing
Set strMsg=Nothing
Set strStatus=Nothing
Set total_time_taken=Nothing
Set threshold_time=Nothing
Set msg=Nothing
Set time_diff=Nothing

Call LogOff()
Call FinalStatus ()


