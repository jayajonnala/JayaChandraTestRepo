'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_PRE_VIM Change CC owner
'.................Test Set Name : AT_04.04.02.47 VIM - PO Precontrole Issue - BR25 - Process PO
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_PRE_VIM Change CC owner"
gstrTCDescription = "PRE_VIM Change CC owner"'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\DLL\FICO\VIM\DT_PRE_VIM Change CC ownerV1.xls"
'gstrresultFolderPath="S:\TASEResult\DLL\FICO"
'datatable_row=157
If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	datatable_row= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'Call SAPGuiUtil.CloseConnections()
'Test initiation
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",157)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'launch SAP

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''''
'
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'-------------------------------------------- KS02----------------------------------------------
Call SetTextbox("Controlling Area","SVALD-VALUE","",DT_KS02_0300_CONTROLLING_AREA,True)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)
Call SetTextbox("Cost Center","CSKSZ-KOSTL","",DT_KS02_0200_COST_CENTER,False)
Call TakeScreenShot()
Call PressEnter()  
Call TakeScreenShot()
Call GetTextboxValue("CSKSZ-VERAK_USER",0,"DT_CHECK_USER_OUTPUT",False)
Call SetTextbox("User Responsible","CSKSZ-VERAK_USER","",DT_KS02_0300_USER_RESPONSIBLE,False)
Call TakeScreenShot()
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call GetTextStatusBar("DT_KS02_0200_GET_TEXT_OF_CHANGES_OUTPUT")
'Call VerifyStatusBar(DT_KS02_0200_CHECK_TEXT_OF_CHANGES)
Call VerifyStatusBar(DT_KS02_0200_GET_TEXT_OF_CHANGES_OUTPUT)

Call LogOff()
Call FinalStatus ()




