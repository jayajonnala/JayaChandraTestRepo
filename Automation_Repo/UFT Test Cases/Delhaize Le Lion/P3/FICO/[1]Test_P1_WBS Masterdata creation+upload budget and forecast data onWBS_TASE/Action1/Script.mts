
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_159_Update Account Completion Table (All Opcos)_TASE
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

'----------------------------------------------------------------------------------------------------------------------------
gstrTestCaseName = "Test_P1_WBS Masterdata creation+upload"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'Login 
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'-------------------------VL10G----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Project def\.","\*PROJ-PSPID","",DT_CJ01_0100_PROJECT_DEF,False)
' SetComboByKey(attachedTextOrComboName, keyValue)
Call SetComboByKey("Project Profile",DT_CJ01_0100_PROJECT_PROFILE)
Call TakeScreenShot
Call ClickButton("Project definition   \(Shift\+F1\)",False)
Call TakeScreenShot
Call SetTextbox("Company code","PROJ-VBUKR","",DT_CJ01_1205_COMPANY_CODE,False)
Call SetTextbox("Profit Center","PROJ-PRCTR","",DT_CJ01_1205_PROFIT_CENTER,False)
Call SetTextbox("Project def\.","PROJ-POST1","",DT_CJ01_0998_PROJECT_DEF,False)
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call VerifyStatusBar(DT_CJ01_0100_CHECK_TEXT_OF_STATUSBAR)

Call LogOff()
Call FinalStatus()

