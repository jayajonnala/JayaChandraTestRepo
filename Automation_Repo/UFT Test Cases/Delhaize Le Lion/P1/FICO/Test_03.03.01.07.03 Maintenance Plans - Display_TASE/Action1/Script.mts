

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_03.03.01.07.03 Maintenance Plans - Display
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

'gstrTestCaseName = "Test_03.03.01.07.03 Maintenance Plans - Display"
gstrTestCaseName = "Test_03.03.01.07.03_Maintenance_Plans_-_Display"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\P1_FICO\DT_03.03.01.07.03 Maintenance Plans - Display_TASE.xls"


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
'''--------------------------------------------  IP03------------------------------------------------
Call SetTextbox("MaintenancePlan","RMIPM-WARPL","",DT_IP03_0140_MAINTENANCEPLAN,False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()

Call VerifyTextBoxContent("Maintenance plan","RMIPM-WARPL",0,DT_IP03_6000_CHECK_TEXT_OF_MAINTENANCE_PLAN,False)
Call VerifyTableCellContent(1,"Cycle","SAPLIWP3TCTRL_0207",DT_IP03_0207_CHECK_TEXT_OF_TABLECELL_CYCLE_0)
Call VerifyTableCellContent(1,"Unit","SAPLIWP3TCTRL_0207",DT_IP03_0207_CHECK_TEXT_OF_TABLECELL_UNIT_0)
Call VerifyTableCellContent(1,"Maintenance cycle text","SAPLIWP3TCTRL_0207",DT_IP03_0207_CHECK_TEXT_OF_TABLECELL_MAINTENANCE_CYCLE_TEXT_0)
Call VerifyTextBoxContent("Functional loc\.","RIWO1-TPLNR",0,DT_IP03_0100_CHECK_TEXT_OF_FUNCTIONAL_LOC,False)
Call VerifyTextBoxContent("Planning site","RMIPM-IWERK",0,DT_IP03_0500_CHECK_TEXT_OF_PLANNING_SITE,False)
Call VerifyTextBoxContent("Order Type","RMIPM-AUART",0,DT_IP03_0500_CHECK_TEXT_OF_ORDER_TYPE,False)
Call VerifyTextBoxNoLabelContent("RMIPM-WERGW",0,DT_IP03_0500_CHECK_TEXT_OF_RMIPMWERGW,False)
Call VerifyTextBoxContent("Main WorkCtr","RMIPM-GEWERK",0,DT_MAINWC,False)
Call VerifyComboBoxValue("Priority" ,DT_IP03_0500_CHECK_VALUE_OF_PRIORITY)
Call VerifyTextBoxContent("Maint\. Planner Group","RMIPM-WPGRP",0,DT_IP03_0500_CHECK_TEXT_OF_MAINT_PLANNER_GROUP,False)
Call VerifyTextBoxContent("Typ","RMIPM-PLNTY",0,DT_IP03_0500_CHECK_TEXT_OF_RMIPMPLNTY,False)
Call VerifyTextBoxContent("Task LstGrp","RMIPM-PLNNR",0,DT_IP03_0500_CHECK_TEXT_OF_RMIPMPLNNR,False)
Call VerifyTextBoxContent("GrpCr","RMIPM-PLNAL",0,DT_IP03_0500_CHECK_TEXT_OF_RMIPMPLNAL,False)

Call SelectTab("TABSTRIP_ITEM","Object list item",False)
Call TakeScreenShot()
Call SelectTab("TABSTRIP_ITEM","Item location",False)
Call TakeScreenShot()
Call VerifyTextBoxContent("Company Code","ILOA-BUKRS",0,DT_COMPANYCODE,False)
Call VerifyTextBoxContent("Business Area","ILOA-GSBER",0,DT_BUSINESSAREA,False)
Call VerifyTextBoxContent("Cost Center","ILOA-KOSTL",0,DT_COSTCENTER,False)
CurrentDateTabName="Cycle item " & ConvertDate(Date)
Call SelectTab("TABSTRIP_ITEM",CurrentDateTabName,False)
Call TakeScreenShot()
Call VerifyTableCellContent(1,"Cycle","SAPLIWP3TCTRL_0208",DT_IP03_0208_CHECK_TEXT_OF_TABLECELL_CYCLE_0)
Call VerifyTableCellContent(1,"Unit","SAPLIWP3TCTRL_0208",DT_IP03_0208_CHECK_TEXT_OF_TABLECELL_UNIT_0)
Call VerifyTableCellContent(1,"Maintenance cycle text","SAPLIWP3TCTRL_0208",DT_IP03_0208_CHECK_TEXT_OF_TABLECELL_MAINTENANCE_CYCLE_TEXT_0)
Call SelectTab("TABSTRIP_HEAD","Maintenance plan scheduling parameters",False)
Call TakeScreenShot()
Call VerifyTextBoxContent("Shift Factor Late Compl\.","RMIPM-VSPOS",0,DT_IP03_0115_CHECK_TEXT_OF_SHIFT_FACTOR_LATE_COMPL,False)
Call VerifyTextBoxContent("Tolerance \(\+\)","RMIPM-TOPOS",0,DT_IP03_0115_CHECK_TEXT_OF_TOLERANCE__,False)
Call VerifyTextBoxContent("Shift Factor Early Compl\.","RMIPM-VSNEG",0,DT_IP03_0115_CHECK_TEXT_OF_SHIFT_FACTOR_EARLY_COMPL,False)
Call VerifyTextBoxContent("Tolerance \(-\)","RMIPM-TONEG",0,DT_IP03_0115_CHECK_TEXT_OF_TOLERANCE___OCC1,False)
Call VerifyTextBoxContent("Call horizon","RMIPM-HORIZ",0,DT_IP03_0115_CHECK_TEXT_OF_CALL_HORIZON,False)
Call VerifyTextBoxContent("Scheduling period","RMIPM-ABRHO",0,DT_IP03_0115_CHECK_TEXT_OF_SCHEDULING_PERIOD,False)
Call VerifyTextBoxContent("Scheduling period","RMIPM-HUNIT",0,DT_IP03_0115_CHECK_TEXT_OF_SCHEDULING_PERIOD_OCC1,False)

Call SelectTab("TABSTRIP_HEAD","Maintenance plan additional data",False)
Call TakeScreenShot()
Call VerifyTextBoxContent("Strategy","RMIPM-WSTRA",0,DT_IP03_0113_CHECK_TEXT_OF_STRATEGY,False)
Call VerifyComboBoxValue("Maint\. plan cat\." ,DT_IP03_0113_CHECK_VALUE_OF_MAINT_PLAN_CAT)

Call LogOff()
Call FinalStatus ()
