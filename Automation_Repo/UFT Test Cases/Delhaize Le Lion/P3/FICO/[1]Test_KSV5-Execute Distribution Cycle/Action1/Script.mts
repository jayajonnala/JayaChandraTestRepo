
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
gstrTestCaseName = "Test_KSV5-Execute Distribution Cycle_TASE"
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

Call SetTextbox("Controlling Area","SVALD-VALUE","",DT_KSV5_0300_CONTROLLING_AREA,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot
Call SetTextbox("Period","RKGA2U-FROM","",DT_KSV5_0101_PERIOD,False)
Call SetTextbox("To","RKGA2U-TO","",DT_KSV5_0101_TO,false)
Call SetTextbox("Fiscal Year","RKGA2U-GJAHR","",DT_KSV5_0101_FISCAL_YEAR,False)
Call TakeScreenShot
Call SetTextbox("Cycle","RKGA2-KSCYC","",DT_KSV5_0101_CYCLE,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
' ClickLabel(labelContent, labelIndex, blnIsItPopup)
Call ClickLabel("Warning","",False)
Call TakeScreenShot
'Call SendKey("{F2}")
Call ClickButton("0 Error  -  Hide   \(Shift\+F5\)",True)
Call TakeScreenShot
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call ClickButton("Yes",True)
Call SelectCheckbox("RKGA2U-TEST","0","OFF",False)
Call ClickButton("Execute   \(F8\)",False)
Call ClickLabel("Warning","",False)
Call ClickButton("0 Error  -  Hide   \(Shift\+F5\)",True)
Call TakeScreenShot
Call ClickButton("Cancel   \(F12\)",True)

Call LogOff()
Call FinalStatus()

