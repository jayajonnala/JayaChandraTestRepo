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

'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'reload DS to update dates and calculations
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Internal order master data maintenance_p4_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 12th July
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Internal order master data maintenance_p4_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\DS\FICO\DT_Internal order master data maintenance_p4_TASE.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
'increment
'Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''
''''----------------------Tcode KOK4----------------------------
''
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Cancel   \(F12\)",True)

Call SelectCheckbox("MKAUF-LISTE","1",DT_KOK4_2010_DETAIL_LIST,False)
Call SetTextbox("Selection variant","CODIA-VARIANT","",DT_KOK4_0510_SELECTION_VARIANT,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Function selection",False)
wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call SelectRowGuiTableByRow("SAPMKAUFTCTRL_MASCHSAMFUNCTSUBS",3,True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Continue   \(Enter\)",True)
wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Execute Selection   \(F8\)",False)
wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Detail list   \(F2\)",False)
wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyGridCellContent("Detail List of Processed Orders",1,"Function","",Lcase(DT_KOK4_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_FUNCTION))
Call VerifyGridCellContent("Detail List of Processed Orders",1,"Status","",Lcase(DT_KOK4_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_STATUS))

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()
