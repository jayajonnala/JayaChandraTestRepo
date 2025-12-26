

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_03.03.01.01.03 Manage Equipment - Change of Functional Location
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

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_03.03.01.01.03 Manage Equipment - Change of Functional Location"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_013_ENA_prices_are_not_higher_than_AB_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''
''----------------------Tcode IE05----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


'Enter Details
Call SetTextbox("Equipment","EQUNR-LOW","",DT_IE05_1000_EQUIPMENT,False) 
Call TakeScreenShot()

'Click on Execute
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()

Call ClickButtonIfExist("Yes",True)
wait(2)

'Navigate to Structure Tab
Call SelectTab("TABSTRIP","Structure",False)
Wait(1)
Call TakeScreenShot()

'Click on Change InstLoc
Call ClickButton("Change InstLoc",False)
Call TakeScreenShot()

'Click on Dismantle
Call ClickButton("Dismantle   \(Shift\+F2\)",True)


Call SetTextbox("Functional loc\.","IEQINSTALL-TPLNR","",DT_IE05_0100_FUNCTIONAL_LOC,True)
Call TakeScreenShot()

'Click on Confirm
Call ClickButton("Confirm   \(Shift\+F4\)",True)

'Navigate to Classification Tab
Call SelectTab("TABSTRIP","Classification",False)
Wait(1)
Call TakeScreenShot()

Call SetTextbox("SERIAL NUMBER","RCTMS-MWERT","1",DT_IE05_4000_RCTMSMWERT,False)
Call TakeScreenShot()

'Navigate to Structure Tab
Call SelectTab("TABSTRIP","Structure",False)
Wait(1)
Call TakeScreenShot()

'Enter Details
Call SetTextbox("TechIdentNo\.","ITOB-TIDNR","",DT_IE05_1065_TECHIDENTNO,False)
Call TakeScreenShot()
Call PressEnter()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call VerifyStatusBarMessageType("S")
Call VerifyStatusBar(DT_SUCCESSMESSAGE)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
