

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_03.03.01.02.04 Manage Function Location - Activate
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

gstrTestCaseName = "Test_03.03.01.02.04 Manage Function Location - Activate"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_013_ENA_prices_are_not_higher_than_AB_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''
''----------------------Tcode IL02----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


'Enter Details
Call SetTextbox("Functional Loc\.","IFLO-TPLNR","",DT_IL02_1110_FUNCTIONAL_LOC,False)
Call SetTextbox("StrIndicator","RILO0-TPLKZ","",DT_IL02_1110_STRINDICATOR,False)
Call TakeScreenShot()
Call PressEnter()  

'Click on Functional location;Functions;Active <-> Inactive;Deactivate
Call SelectMenuBar("Functional location;Functions;Active <-> Inactive;Deactivate")
Call TakeScreenShot()

'save The details
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call VerifyStatusBarMessageType("S")

Call GetStatusBar("item1","DT_FUNCLOC_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar("Functional location "&DT_IL02_1110_CHECK_MESSAGEPARAMETER_OF_STATUSBAR&" changed" )

'Click on Functional location;Display
Call SelectMenuBar("Functional location;Display")
Call TakeScreenShot()

Call PressEnter()

Call VerifyStatusBarMessageType("S")

Call GetTextboxValue("RILO0-STTXT",0,"DT_STATUS_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyTextBoxContent("Status","RILO0-STTXT",0,DT_IL02_2100_CHECK_TEXT_OF_STATUS_OCC1,False)

'Click on Functional location;Display -> Change
Call SelectMenuBar("Functional location;Display -> Change")
Call TakeScreenShot()

Call ClickButtonIfExist("Yes",True)
wait(2)
Call TakeScreenShot()


'Click on Functional location;Display -> Change
Call SelectMenuBar("Functional location;Functions;Active <-> Inactive;Activate")
Call TakeScreenShot()

'Save The Details
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call VerifyStatusBarMessageType("S")


Call GetStatusBar("item1","DT_FUNCLOC_OUTPUT1")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar("Functional location "&DT_IL02_1110_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC2&" changed" )
Call PressEnter() 

Call VerifyTextBoxContent("Status","RILO0-STTXT",0,DT_IL02_2100_CHECK_TEXT_OF_STATUS_OCC2,False)


'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

