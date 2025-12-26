

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_04.01.05.07.01 Execute Preventive Maintenance (Internal)
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

'gstrTestCaseName = "Test_04.01.05.07.01 Execute Preventive Maintenance (Internal)"
gstrTestCaseName = "Test_04.01.05.07.01_Execute_Preventive_Maintenance_(Internal)"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\P1_FICO\DT_04.01.05.07.01 Execute Preventive Maintenance (Internal)_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'
'''--------------------------------------------  IP17------------------------------------------------
Call SetTextbox("Maintenance Plan","WARPL-LOW","",DT_IP17_1000_MAINTENANCE_PLAN,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()
Call SelectRowGuiGrid("",0,"Maintenance Plan",DT_IP17_1000_MAINTENANCE_PLAN,False)
Call DoubleClick()
Call TakeScreenShot()
Call SelectTab("TABSTRIP_ITEM","Object list item",False)
Call TakeScreenShot()
Call SelectTab("TABSTRIP_ITEM","Item location",False)
Call TakeScreenShot()
Call ClickButtonIfExist("Back   \(F3\)",False)
Call TakeScreenShot()
Call ClickButtonIfExist("No",True)
Call PressEnter() 
Call TakeScreenShot()
Call SelectRowGuiGrid("",0,"Maintenance Plan",DT_IP17_1000_MAINTENANCE_PLAN,False)
Call ClickButtonIfExist("Maintenance plans   \(Ctrl\+Shift\+F8\)",False)
Call TakeScreenShot()
Call SelectRowGuiGrid("",0,"Maintenance Plan",DT_IP17_1000_MAINTENANCE_PLAN,False)
Call DoubleClick()
Call TakeScreenShot()
Call SelectTab("TABSTRIP_ITEM","Schedule call item",False)
Call VerifyTableCellContent(1,"PlanDate","SAPLIWP3TCTRL_0123",ConvertDate(DT_IP17_0123_CHECK_TEXT_OF_TABLECELL_PLANDATE_0))
Call VerifyTableCellContent(1,"Due packages","SAPLIWP3TCTRL_0123",DT_IP17_0123_CHECK_TEXT_OF_TABLECELL_DUE_PACKAGES_0)
Call VerifyTableCellContent(1,"Scheduling Type / Status","SAPLIWP3TCTRL_0123",LCASE(DT_IP17_0123_CHECK_TEXT_OF_TABLECELL_SCHEDULING_TYPE__STATUS_0))
Call VerifyTableCellContent(1,"Call date","SAPLIWP3TCTRL_0123",ConvertDate(DT_IP17_0123_CHECK_TEXT_OF_TABLECELL_CALL_DATE_0))


'''--------------------------------------------  IP10------------------------------------------------
'---
Call SetTcode(DT_IP17_0201_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_IP17_0201_OKCD)

Call PressEnter()
Call TakeScreenShot()

Call ClickButton("New Start   \(Shift\+F7\)",False)
Call TakeScreenShot()
Call ClickButtonIfExist("Delete",True)
Call SetTextbox("Start of cycle","RMIPM-STADT","",ConvertDate(DT_IP17_7000_START_OF_CYCLE),True)
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_IP17_0140_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Maintenance plan "& DT_IP17_0140_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT &" scheduled" )

'''--------------------------------------------  IP17------------------------------------------------
'---
Call SetTcode(DT_IP17_0140_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_IP17_0140_OKCD)

Call SetTextbox("Maintenance Plan","WARPL-LOW","",DT_IP17_1000_MAINTENANCE_PLAN,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()
Call SelectRowGuiGrid("",0,"Maintenance Plan",DT_IP17_1000_MAINTENANCE_PLAN,False)
Call ClickButtonIfExist("Maintenance plans   \(Ctrl\+Shift\+F8\)",False)
Call TakeScreenShot()
Call ClickButtonIfExist("Maintenance plan   \(Ctrl\+Shift\+F7\)",False)
Call SelectTab("TABSTRIP_ITEM","Schedule call item",False)
Call VerifyTableCellContent(1,"PlanDate","SAPLIWP3TCTRL_0123",ConvertDate(DT_IP17_0123_CHECK_TEXT_OF_TABLECELL_PLANDATE_0_OCC1))
Call VerifyTableCellContent(1,"Due packages","SAPLIWP3TCTRL_0123",DT_IP17_0123_CHECK_TEXT_OF_TABLECELL_DUE_PACKAGES_0_OCC1)
Call VerifyTableCellContent(1,"Scheduling Type / Status","SAPLIWP3TCTRL_0123",LCASE(DT_IP17_0123_CHECK_TEXT_SCHEDULING_TYPE__STATUS_0_OCC1))
Call VerifyTableCellContent(1,"Call date","SAPLIWP3TCTRL_0123",ConvertDate(DT_IP17_0123_CHECK_TEXT_OF_TABLECELL_CALL_DATE_0_OCC1))
'
Call LogOff()
Call FinalStatus ()

