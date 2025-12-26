'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Depreciate Assets_p4_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 19th March
'.................Modified By :
'.................Modified Date/Details :
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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

gstrTestCaseName = "Test_Depreciate Assets_p4_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Fixed Assets Sales_p3_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call TakeScreenShot()

Call SetTextbox("Company Code","P_BUKRS","",DT_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","P_GJAHR","",DT_YEAR,False)
Call SetTextbox("Posting Period","P_BUPER","",DT_POSTING_PERIOD,False)
Call SelectCheckBox("P_SCHED",0,"ON",False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()
Call ActivateNodeGuiTree(0,"#1")
Call TakeScreenShot()
Call VerifyGridCellContent("",1,"Message Type",0,DT_STATUS)

Call SetTcode(DT_SAPTCODE_OCC1) 
Call PressEnter()     ' 
Call TakeScreenShot()
Call ClickButton("All Selections   \(Shift\+F7\)",False)
Call SetTextbox("Company code","BUKRS-LOW","",DT_COMPANY_CODE,False)
Call SetTextbox("Asset number","ANLAGE-LOW","",DT_ASSET,False)
Call SetTextbox("Report date","BERDATUM","",DT_REPORT_DATE,False)
Call SetTextbox("Depreciation area","BEREICH1","",DT_DEPRECIATION_AREA,False)
Call SetTextbox("Sort Variant","SRTVR","",DT_SORT_VARIANT,False)
Call SelectRadioButton("XEINZEL","List assets",False)
Call SelectRadioButton("PA_MONTH","Month",False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()


Call LogOff()
Call FinalStatus ()




