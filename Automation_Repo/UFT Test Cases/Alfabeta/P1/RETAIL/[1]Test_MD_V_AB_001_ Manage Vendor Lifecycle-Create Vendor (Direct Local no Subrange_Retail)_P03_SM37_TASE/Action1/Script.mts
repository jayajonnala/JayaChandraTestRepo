

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_V_AB_001_ Manage Vendor Lifecycle-Create Vendor (Direct Local no Subrange_Retail)_P03_SM37
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

gstrTestCaseName = "Test_MD_V_AB_001_P03_SM37"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\P1_FICO\DT_04.04.02.21 VIM - PO Precontrole Issue - BR01 - Invalid Company_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

'DataRowSet =3
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE_OCC1)

''--------------------------------------------  SM37------------------------------------------------
Call SelectCheckbox("BTCH2170-PRELIM",0,DT_SM37_2170_SCHED,False)
Call SetTextbox("User Name","BTCH2170-USERNAME","",DT_USERNAME,False)
Call ClickButton("Execute   \(F8\)",False)
Call ClickButton("btn\[8\]",False)
Call TakeScreenShot()

''----------------------------
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
'''Call SelectRowGuiTableByRow("SAPLSKBHTC_FIELD_LIST_820",5,True)

Call ClickButtonToolBar("&FIND",0)
'''Call ClickButton("&FIND",True)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Status",True)
Call TakeScreenShot()
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call TakeScreenShot()

'''Call ClickButton("Show sel\. fields \(CTRL\+F3\)",True)
Call ClickButton("Add Filter Criterion \(F7\)",True)
Call ClickButton("Define Filter Values",True)
Call TakeScreenShot()

'''Call ClickButton("Copy   \(Enter\)",True)
Call SetTextbox("Status","%%DYN001-LOW","",DT_SM37_1105_STATUS,True)
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot()

'Call VerifyifGuiLabelExists_ByIndex(DT_SM37_1105_STATUS,0)
'Call GetLabelContentByRefLabel("JobName",0,-32,"DT_SM37_0120_CHECK_TEXT_OF_LOCL_OUTPUT",False)
'Call GetLabelContentByRefLabel("Start Time",0,-32,"DT_SM37_START_TIME_OUTPUT",False)
CAll SelectColumnGuiGrid("", "", "Start Time", False)
Call ClickButton("Sort in Descending Order   \(Ctrl\+Shift\+F4\)",False)

Call VerifyGridCellContent("", 1, "Status", "", Lcase(DT_SM37_1105_STATUS))
Call GetGridContentByTitle("", "", "JobName", 1, "DT_SM37_0120_CHECK_TEXT_OF_LOCL_OUTPUT")
Call GetGridContentByTitle("","","Start Time",1,"DT_SM37_START_TIME_OUTPUT")

Call LogOff()
Call FinalStatus ()

