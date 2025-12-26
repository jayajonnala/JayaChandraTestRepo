
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



gstrTestCaseName = "Test_Post GI"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'--------------------------------------VL02N--------------------
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


Call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_VL02N_4004_OUTBOUND_DELIVERY,False)  
Call TakeScreenShot()
call ClickButton("Post Goods Issue   \(Shift\+F8\)",false)

If VerifyStatusBarConditonNoReport(DT_VL02N_4004_CHECK_TEXT_OF_STATUSBAR) <> True then
     Call VerifyGridCellContent("",1,"Message Text",0,DT_VL02N_4004_CHECK_TEXT_OF_STATUSBAR_OCC1)
     Call TakeScreenShot
Else
 Call VerifyStatusBar(DT_VL02N_4004_CHECK_TEXT_OF_STATUSBAR)
     Call TakeScreenShot
     
End If
'Call TakeScreenShot()
'Call GetStatusBar("text","DT_Statusbar_output")
'call VerifyStatusBar(DT_Statusbar_output)
'Call TakeScreenShot()


Call LogOff()
Call FinalStatus ()

