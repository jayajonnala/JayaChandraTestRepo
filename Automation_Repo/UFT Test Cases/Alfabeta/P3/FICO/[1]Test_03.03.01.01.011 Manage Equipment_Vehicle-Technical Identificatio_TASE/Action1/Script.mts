

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_03.03.01.01.011 Manage Equipment_Vehicle-Technical Identificatio
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

gstrTestCaseName = "Test_03.03.01.01.011 Manage Equipment_Vehicle-Technical Identificatio"
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

'Enter Details
Call SetTextbox("TechIdentNo\.","ITOB-TIDNR","",DT_IE05_1065_TECHIDENTNO,False)
Call TakeScreenShot()

' ClickButton(tooltipOrButtonName, blnIsItPopup)
Call ClickButton("Save   \(Ctrl\+S\)",False)

Call VerifyTextBoxContent("Information Message","MESSTXT1",0,Ucase(DT_IE05_0010_CHECK_TEXT_OF_MESSTXT1),True)

Call ClickButtonIfExist("Continue   \(Enter\)",True)
wait(2)

Call ClickButtonIfExist("Cancel   \(F12\)",False)
wait(2)

Call ClickButtonIfExist("Yes",True)
wait(2)


'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************


