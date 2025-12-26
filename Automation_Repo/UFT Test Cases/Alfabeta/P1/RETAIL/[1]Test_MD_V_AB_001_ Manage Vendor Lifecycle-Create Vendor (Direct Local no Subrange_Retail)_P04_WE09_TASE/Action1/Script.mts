

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_V_AB_001_ Manage Vendor Lifecycle-Create Vendor (Direct Local no Subrange_Retail)_P04_WE09
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
	GetRowNo= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrTestCaseName = "Test_MD_V_AB_001_ Manage Vendor_P04_WE09"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'
'
''''''--------------------------------WE09-----------------------------
'
 Call SetTextbox("Created At","CRETIM-LOW","",DT_WE09_1000_CREATED_AT,False)

 Call SetTextbox("Logical Message","MESTYP-LOW","",DT_WE09_1000_LOGICAL_MESSAGE,False)
 Call SetTextbox("Search in Segment \.\.\.","SEGMENT1","",DT_WE09_1000_SEARCH_IN_SEGMENT_,False)
 Call SetTextbox("Search in Field \.\.\.","FIELD1_1","",DT_WE09_1000_SEARCH_IN_FIELD_,False) 
  Call SetTextbox("for Value ...","VALUE1_1",0,DT_WE09_1000_FOR_VALUE_,False)
 Call TakeScreenShot()
  Call PressEnter()  ' 
   Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False) 
Call VerifyifGuiLabelExists_ByIndex(DT_WE09_0120_CHECK_TEXT_OF_DATA_PASSED_TO_PORT_OK,0)
Call VerifyifGuiLabelExists_ByIndex(DT_WE09_0120_CHECK_TEXT_OF_NO_NAME,0)

Call GetLabelContentByRefLabel("IDoc number",0,-48,"DT_IDOC_OUTPUT",False)
Call ClickLabel(DT_IDOC_OUTPUT,0,False)
Call TakeScreenShot()
Call ClickButton("Back   \(F3\)",False) 
Call TakeScreenShot()



Call LogOff()
Call FinalStatus ()






