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
'.................Test Script Name : Test_Internal order master data maintenance_p3_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 12th July
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Internal order master data maintenance_p3_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\DS\FICO\DT_Internal order master data maintenance_P3_TASE.xls"
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
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''
''''----------------------Tcode KO03----------------------------
''
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Order","COAS-AUFNR","",DT_KO03_0110_ORDER,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()   
'Capture the screenshot
Call TakeScreenShot()
Call VerifyTextBoxContent("Order","COAS-AUFNR","",DT_KO03_0600_CHECK_TEXT_OF_ORDER,False)

Call SelectTab("TABSTRIP_600",DT_KO03_0600_CONTROL_DATA,False)
Call VerifyTextBoxContent("Actual posted CCtr","COAS-CYCLE","",DT_KO03_0325_CHECK_TEXT_OF_ACTUAL_POSTED_CCTR,False)
'Capture the screenshot
Call TakeScreenShot()

''
''''----------------------Tcode KOK2----------------------------
''
'Enter the Tcode
Call SetTcode(DT_KO03_0600_OKCD) 
Call PressEnter()     ' 
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Cancel   \(F12\)",True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Create selection variant",False)

Call SetTextbox("Variant","RSVAR-VARIANT","",DT_KO03_0302_VARIANT,False)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButtonIfExist("Create",False)

Call SetTextbox("Order type","AUART-LOW","",DT_KO03_1000_ORDER_TYPE,False)
Call SetTextbox("Company code","BUKRS-LOW","",DT_KO03_1000_COMPANY_CODE,False)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButtonIfExist("Attributes   \(F6\)",False)

Call SetTextbox("Description","RSVAR-VTEXT","",DT_KO03_0281_DESCRIPTION,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()
'veryfy sattus bar content
Call GetStatusBar("item1","DT_KO03_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_KO03_1000_CHECK_TEXT_OF_STATUSBAR)

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()
