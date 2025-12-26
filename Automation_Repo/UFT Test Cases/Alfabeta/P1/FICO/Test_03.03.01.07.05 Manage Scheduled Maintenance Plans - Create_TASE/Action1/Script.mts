

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_03.03.01.07.05 Manage Scheduled Maintenance Plans - Create
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

'gstrTestCaseName = "Test_03.03.01.07.05 Manage Scheduled Maintenance Plans - Create"
gstrTestCaseName = "Test_03.03.01.07.05_Manage_Scheduled_Maintenance_Plans_-_Create"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\P1_FICO\DT_03.03.01.07.05 Manage Scheduled Maintenance Plans - Create_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
'''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'
'''--------------------------------------------  IP10------------------------------------------------
Call SetTextbox("MaintenancePlan","RMIPM-WARPL","",DT_IP10_0140_MAINTENANCEPLAN,False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()
Call ClickButton("Start   \(F9\)",False)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call VerifyTableCellContent(1,"PlanDate","SAPLIWP3TCTRL_0121",ConvertDate(DT_IP10_0121_CHECK_TEXT_OF_TABLECELL_CALL_DATE_0))
Call VerifyTableCellContent(1,"Due packages","SAPLIWP3TCTRL_0121",DT_IP10_0121_CHECK_TEXT_OF_TABLECELL_DUE_PACKAGES_0)
Call VerifyTableCellContent(1,"Scheduling Type / Status","SAPLIWP3TCTRL_0121",LCASE(DT_IP10_0121_CHECK_TEXT_OF_TABLECELL_SCHEDULING_TYPE__STATUS_0))
Call VerifyTableCellContent(1,"Call date","SAPLIWP3TCTRL_0121",ConvertDate(DT_IP10_0121_CHECK_TEXT_OF_TABLECELL_PLANDATE_0))


Call ClickButton("Save   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_IP10_0140_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Maintenance plan "& DT_IP10_0140_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT &" scheduled" )


Call LogOff()
Call FinalStatus ()




