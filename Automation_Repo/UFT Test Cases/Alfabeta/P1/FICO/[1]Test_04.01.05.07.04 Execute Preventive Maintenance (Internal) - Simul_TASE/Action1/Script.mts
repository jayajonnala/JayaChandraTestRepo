

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_04.01.05.07.04 Execute Preventive Maintenance (Internal) - Simul
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

'gstrTestCaseName = "Test_04.01.05.07.04 Execute Preventive Maintenance (Internal) - Simul"
gstrTestCaseName = "Test_04.01.05.07.04_Execute_Preventive_Maintenance_(Internal)_-_Simul"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\P1_FICO\DT_04.01.05.07.04 Execute Preventive Maintenance (Internal) - Simul_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
''
'''''--------------------------------------------  IP19------------------------------------------------
Call SelectRadioButton("GRAPH3","Tabular with Navigation Tree",False)
Call SetTextbox("Maintenance Plan","WARPL-LOW","",DT_IP19_1000_MAINTENANCE_PLAN,False)
Call SetTextbox("From Date","DATUV","",DT_IP19_1000_FROM_DATE,False)
Call SetTextbox("To Date","DATUB","",DT_IP19_1000_TO_DATE,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call ActivateNodeGuiTree(0,"#1;#1")
Call ClickButton("Simulate Maintenance Plan   \(Shift\+F7\)",False)
Call ActivateNodeGuiTree(0,"#1;#1")

Call VerifyGridCellContent("Maint\. Dates",1,"WARPL",0,DT_IP19_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WARPL)
Call VerifyGridCellContent("Maint\. Dates",1,"CALL_STATUS",0,UCASE(DT_IP19_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_CALL_STATUS))
Call VerifyGridCellContent("Maint\. Dates",1,"NPLDA",0,ConvertDate(DT_IP19_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_NPLDA))

Call ClickButton("Release Call   \(Shift\+F4\)",False)
wait 5
Call VerifyGridCellContent("Maint\. Dates",1,"CALL_STATUS",0,UCASE(DT_IP19_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_CALL_STATUS_OCC1))

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
wait 5
Call GetStatusBar("item1","DT_IP19_0200_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Order saved with number " & DT_IP19_0200_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT)
IP19_0101_ORDER=DT_IP19_0200_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT

'''--------------------------------------------  IW33------------------------------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_IP19_0200_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_IP19_0200_OKCD)

Call SetTextbox("Order","CAUFVD-AUFNR","",IP19_0101_ORDER,False)
Call TakeScreenShot()
Call PressEnter()     ' 
Call TakeScreenShot()

Call SelectTab("TS_1100","Location",False)
Call TakeScreenShot()
Call VerifyTextBoxContent("MaintSite","ILOA-SWERK",0,UCASE(DT_IP19_7000_CHECK_TEXT_OF_MAINTSITE),False)
Call VerifyTextBoxContent("Company Code","ILOA-BUKRS",0,UCASE(DT_IP19_7000_CHECK_TEXT_OF_COMPANY_CODE),False)
Call VerifyTextBoxContent("Business Area","ILOA-GSBER",0,UCASE(DT_IP19_7000_CHECK_TEXT_OF_BUSINESS_AREA),False)
Call VerifyTextBoxContent("Cost Center","ILOA-KOSTL",0,UCASE(DT_IP19_7000_CHECK_TEXT_OF_COST_CENTER),False)
Call VerifyTextBoxContent("CO Area","ILOA-KOKRS",0,UCASE(DT_IP19_7000_CHECK_TEXT_OF_CO_AREA),False)
  
Call SelectTab("TS_1100","Planning",False)
Call TakeScreenShot()
Call VerifyTextBoxContent("MaintenancePlan","CAUFVD-WARPL",0,DT_IP19_1160_CHECK_TEXT_OF_MAINTENANCEPLAN,False)
  
Call SelectTab("TS_1100","Control",False)
Call TakeScreenShot()
Call VerifyTextBoxContent("Entered by","CAUFVD-ERNAM",0,UCASE(DT_IP19_1140_CHECK_TEXT_OF_ENTERED_BY),False)
Call VerifyTextBoxContent("Created on","CAUFVD-ERDAT",0,ConvertDate(DT_IP19_1140_CHECK_TEXT_OF_CREATED_ON),False)
Call VerifyTextBoxContent("CstgVariantPlan","CAUFVD-KLVARP",0,UCASE(DT_IP19_1140_CHECK_TEXT_OF_CSTGVARIANTPLAN),False)
Call VerifyTextBoxContent("CstgVariantActl","CAUFVD-KLVARI",0,UCASE(DT_IP19_1140_CHECK_TEXT_OF_CSTGVARIANTACTL),False)
Call VerifyTextBoxContent("Status Profile","CAUFVD-STATS",0,UCASE(DT_IP19_1140_CHECK_TEXT_OF_STATUS_PROFILE),False)
Call VerifyComboBoxValue("Res\./Purc\. req\.",DT_IP19_1141_CHECK_VALUE_OF_RESPURC_REQ)
 
Call SelectTab("TS_1100","Additional Data",False)
Call TakeScreenShot()
Call VerifyTextBoxContent("Company Code","CAUFVD-BUKRS",0,UCASE(DT_IP19_1130_CHECK_TEXT_OF_COMPANY_CODE),False)
Call VerifyTextBoxContent("Business Area","CAUFVD-GSBER",0,UCASE(DT_IP19_1130_CHECK_TEXT_OF_BUSINESS_AREA),False)
Call VerifyTextBoxContent("Controlling Area","CAUFVD-KOKRS",0,UCASE(DT_IP19_1130_CHECK_TEXT_OF_CONTROLLING_AREA),False)
Call VerifyTextBoxContent("Responsible CCtr","CAUFVD-KOSTV",0,UCASE(DT_IP19_1130_CHECK_TEXT_OF_RESPONSIBLE_CCTR),False)
Call VerifyTextBoxContent("Profit Center","CAUFVD-PRCTR",0,UCASE(DT_IP19_1130_CHECK_TEXT_OF_PROFIT_CENTER),False)
Call VerifyComboBoxValue("Object Class",DT_IP19_1130_CHECK_VALUE_OF_OBJECT_CLASS)
'Call VerifyTextBoxContent("Functional Area","CAUFVD-FUNC_AREA",0,UCASE(DT_IP19_1130_CHECK_TEXT_OF_FUNCTIONAL_AREA),False)
Call VerifyTextBoxContent("Processing group","CAUFVD-ABKRS",0,DT_IP19_1130_CHECK_TEXT_OF_PROCESSING_GROUP,False)

 
Call LogOff()
Call FinalStatus ()


