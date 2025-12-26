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
'.................Test Script Name : Test_Cost element group maintenance_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 13th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Cost element group maintenance_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Cost element group maintenance_TASE.xls"
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
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_1",(Cint(DT_INCREMENT_1)+1))
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_2",(Cint(DT_INCREMENT_2)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'''----------------------Tcode KAH1----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Cost element group","GRPDYNP-NAME_COALL","",DT_KAH1_1000_COST_ELEMENT_GROUP,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox(DT_KAH1_1000_COST_ELEMENT_GROUP,"","0",DT_KAH1_0120_NO_NAME,False)
Call ClickLabel(DT_KAH1_1000_COST_ELEMENT_GROUP,"",False)
Call ClickButton("Insert Cost Element   \(Shift\+F4\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox(DT_KAH1_1000_COST_ELEMENT_GROUP,"","6",DT_KAH1_0120_NO_NAME_OCC1,False)
Call SetTextbox(DT_KAH1_1000_COST_ELEMENT_GROUP,"","7",DT_KAH1_0120_NO_NAME_OCC2,False)

Call ClickButton("Confirm Changes   \(Enter\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickLabel(DT_KAH1_1000_COST_ELEMENT_GROUP,"",False)
Call ClickButton("Insert Cost element group on lower level   \(Ctrl\+F3\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call SendKey(DT_KAH1_0120_NO_NAME_OCC3)
Wait(5)
Call SendKey("{TAB}")
Wait(2)
Call SendKey(DT_KAH1_0120_NO_NAME_OCC4)
Wait(7)

Call ClickButton("Confirm Changes   \(Enter\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
'Capture the screenshot
Call TakeScreenShot()
'
'''----------------------Tcode KAH2----------------------------
'
'Enter the Tcode
Call SetTcode(DT_KAH1_0120_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_KAH1_0120_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Cost element group","GRPDYNP-NAME_COALL","",DT_KAH1_1000_COST_ELEMENT_GROUP_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call SetFocusGuiLabel(DT_KAH1_0120_NO_NAME_OCC3,"","",False)

Call ClickButton("Select   \(F9\)",False)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Delete   \(F6\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Remove",True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
'Capture the screenshot
Call TakeScreenShot()
'
'''----------------------Tcode KAH3----------------------------
'
'Enter the Tcode
Call SetTcode(DT_KAH1_0120_OKCD_OCC1) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_KAH1_0120_OKCD_OCC1)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Cost element group","GRPDYNP-NAME_COALL","",DT_KAH1_1000_COST_ELEMENT_GROUP_OCC2,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call VerifyifGuiLabelExists(DT_KAH1_0120_CHECK_TEXT_OF_NO_NAME)
Call VerifyifGuiLabelExists(DT_KAH1_0120_CHECK_TEXT_OF_NO_NAME_OCC1)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

