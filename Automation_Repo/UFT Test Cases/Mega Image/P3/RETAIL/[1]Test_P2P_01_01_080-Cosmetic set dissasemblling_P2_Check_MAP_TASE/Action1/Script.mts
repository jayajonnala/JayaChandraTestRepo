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
'.................Test Script Name : Test_P2P_01_01_080-Cosmetic set dissasemblling_P2_Check_MAP_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 7th June
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_P2P_01_01_080_P2_Check_MAP_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_P2P_01_01_080-Cosmetic set dissasemblling_P2_Check_MAP_TASE.xls"
'''''----------------------Login----------------------------

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

'''----------------------Tcode SM37----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
'Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Job Name","BTCH2170-JOBNAME","",DT_MIGO_2170_JOB_NAME,False)
Call SetTextbox("User Name","BTCH2170-USERNAME","",DT_MIGO_2170_USER_NAME,False)
Call SelectCheckbox("BTCH2170-PRELIM",0,DT_MIGO_2170_SCHED,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
'Capture the screenshot
Call TakeScreenShot()
''
''''----------------------Tcode ZMDIM_MBMAPCHANGES----------------------------
'Enter the Tcode
Call SetTcode(DT_MIGO_3000_OKCD_OCC4) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_MIGO_3000_OKCD_OCC4)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Material","P_MATNR","",DT_ZMDIM_MBMAPCHANGES_1000_MATERIAL,False)
Call SetTextbox("Valuation Area","P_BWKEY","",DT_ZMDIM_MBMAPCHANGES_1000_VALUATION_AREA,False)
Call SetTextbox("Reference Document","P_DOCU","",DT_ZMDIM_MBMAPCHANGES_1000_REFERENCE_DOCUMENT,False)
Call SetTextbox("Year","P_YEAR","1",DT_ZMDIM_MBMAPCHANGES_1000_YEAR,False)
SAPGuiSession("transaction:=ZMDIM_MBMAPCHANGES").SAPGuiWindow("transaction:=ZMDIM_MBMAPCHANGES").SAPGuiRadioButton("name:=P_FROM").Set

'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()
Wait(2)

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()


