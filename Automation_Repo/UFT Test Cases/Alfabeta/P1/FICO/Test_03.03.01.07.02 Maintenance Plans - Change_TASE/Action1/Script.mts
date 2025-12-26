

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_03.03.01.07.02 Maintenance Plans - Change
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

'gstrTestCaseName = "Test_03.03.01.07.02 Maintenance Plans - Change"
gstrTestCaseName = "Test_03.03.01.07.02_Maintenance_Plans_-_Change"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\P1_FICO\DT_03.03.01.07.02 Maintenance Plans - Change_TASE.xls"


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
'''--------------------------------------------  IP02------------------------------------------------
Call SetTextbox("MaintenancePlan","RMIPM-WARPL","",DT_IP02_0140_MAINTENANCEPLAN,False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()
Call SetTextbox("Maintenance plan","RMIPM-WPTXT","",DT_PLANNAMEMOD,False)
Call SelectTab("TABSTRIP_HEAD","Maintenance plan scheduling parameters",False)
Call TakeScreenShot()
Call SelectTab("TABSTRIP_HEAD","Maintenance plan additional data",False)
Call TakeScreenShot()
Call SelectTab("TABSTRIP_HEAD","Maintenance plan scheduling parameters",False)
Call SetTextbox("Call horizon","RMIPM-HORIZ","",DT_IP02_0115_CALL_HORIZON,False)
Call GetTextboxValue("RMIPM-WAPOS",0,"DT_MAINTENANCE_ITEM_OUTPUT",False)

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call VerifyStatusBarMessageType(DT_IP02_0140_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetStatusBar("item1","DT_MAINTENANCEPLAN_UPDATED_OUTPUT")
Call VerifyStatusBar("Maintenance plan "& DT_MAINTENANCEPLAN_UPDATED_OUTPUT&" changed" )


Call LogOff()
Call FinalStatus()

