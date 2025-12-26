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


'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Selection variant maintenance - Cost element group_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 12th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Selection variant maintenance - Cost element group_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Selection variant maintenance - Cost element group_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'Increment the parameter
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'''----------------------Tcode KM5V----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Create",False)
Call SetTextbox("Create variant","RSVAR-VARIANT","",DT_KM5V_0304_CREATE_VARIANT,True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Create   \(Enter\)",True)
'Capture the screenshot
Call TakeScreenShot()

Call FocusTextBox("CElem category","KATYP-LOW",False)
'Call SendKey("{F4}")
'Capture the screenshot
Call TakeScreenShot()

Call ClickLabel("Primary costs/cost-reducing revenues","",True)
Call ClickButtonIfExist("Copy   \(Enter\)",True)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Cost Center","KOSTL-LOW","",DT_KM5V_1000_COST_CENTER,False)
Call FocusTextBox("Cost Center","KOSTL-LOW",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Attributes   \(F6\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Description","RSVAR-VTEXT","",DT_KM5V_0281_DESCRIPTION,False)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)

wait(1)
Call TakeScreenShot()

Call GetStatusBar("item1","DT_KM5V_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

VerifyStatusBar(DT_KM5V_1000_CHECK_TEXT_OF_STATUSBAR)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)

wait(1)
Call TakeScreenShot()

Call GetStatusBar("item1","DT_KM5V_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

VerifyStatusBar(DT_KM5V_1000_CHECK_TEXT_OF_STATUSBAR_OCC1)
'Capture the screenshot
Call TakeScreenShot()


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

