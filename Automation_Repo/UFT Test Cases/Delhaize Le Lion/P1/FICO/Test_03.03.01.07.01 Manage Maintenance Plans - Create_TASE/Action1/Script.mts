

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_03.03.01.07.01 Manage Maintenance Plans - Create
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



gstrTestCaseName = "Test_03.03.01.07.01_Manage_Maintenance_Plans_-_Create"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\P1_FICO\DT_03.03.01.07.01 Manage Maintenance Plans - Create_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
'''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'
'''--------------------------------------------  IP42------------------------------------------------
Call SetComboByKey("RMIPM-MPTYP",DT_IP42_0100_MAINT_PLAN_CAT)
Call SetTextbox("Strategy","RMIPM-WSTRA","",DT_IP42_0100_STRATEGY,False)
Call TakeScreenShot()
Call PressEnter()

Call SetTextbox("Maintenance plan","RMIPM-WPTXT","",DT_IP42_6000_MAINTENANCE_PLAN,False)
Call TakeScreenShot()
Call SetTextbox("Functional loc\.","RIWO1-TPLNR","",DT_IP42_0100_FUNCTIONAL_LOC,False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()
Call SetTextbox("Order Type","RMIPM-AUART","",DT_IP42_0500_ORDER_TYPE,False)
Call SetTextbox("Main WorkCtr","RMIPM-GEWERK","",DT_IP42_0500_MAIN_WORKCTR,False)
Call TakeScreenShot()
Call PressEnter()
Call SetTextbox("Maint\. Planner Group","RMIPM-WPGRP","",DT_IP42_0500_MAINT_PLANNER_GROUP,False)
Call SetComboByKey("RMIPM-PRIOK",DT_IP42_0500_PRIORITY)
Call TakeScreenShot()

Call ClickButton("Select task list",False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()
Call SelectRowGuiGrid("",0,"Group","M061",False)
Call DoubleClick()
''Call ClickButton("Choose   \(Ctrl\+Shift\+F11\)",False)
Call TakeScreenShot()

Call SelectTab("TABSTRIP_HEAD","Maintenance plan scheduling parameters",False)
Call TakeScreenShot()

Call SetTextbox("Tolerance \(-\)","RMIPM-TONEG","",DT_IP42_0115_TOLERANCE___OCC1,False)
Call SetTextbox("Shift Factor Early Compl\.","RMIPM-VSNEG","",DT_IP42_0115_SHIFT_FACTOR_EARLY_COMPL,False)
Call SetTextbox("Scheduling period","RMIPM-HUNIT","",DT_IP42_0115_SCHEDULING_PERIOD_OCC1,False)
Call SetTextbox("Scheduling period","RMIPM-ABRHO","",DT_IP42_0115_SCHEDULING_PERIOD,False)
Call SetTextbox("Call horizon","RMIPM-HORIZ","",DT_IP42_0115_CALL_HORIZON,False)
Call SetTextbox("Tolerance \(\+\)","RMIPM-TOPOS","",DT_IP42_0115_TOLERANCE,False)
Call SetTextbox("Shift Factor Late Compl\.","RMIPM-VSPOS","",DT_IP42_0115_SHIFT_FACTOR_LATE_COMPL,False)
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)

Call GetStatusBar("item1","DT_MAINTENANCEPLAN_OUTPUT")
'Call VerifyStatusBar("Maintenance plan "&DT_MAINTENANCEPLAN_OUTPUT&" created" )


Call LogOff()
Call FinalStatus ()

