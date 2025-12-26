
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Distribution process to DS stores - SW41 (Fresh meat) - p4
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

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Distribution process to DS stores - SW41 (Fresh meat) - p4
'.................Author : TCS 	   :Raushan
'................ Creation Date    :12th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Distribution process to DS stores - SW59 (Dry goods)_p4"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Distribution process to DS stores - SW41 (Fresh meat) - p4.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode VL09N----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Inbound / Outbound Delivery","I_VBELN-LOW","",DT_VL09_1000_INBOUND__OUTBOUND_DELIVERY,FALSE)
Call SelectRadioButton("LF_ANAUS","Inbound Delivs & Outbound Delivs",False)
Call TakeScreenShot()

'Click on Execute
Call ClickButton("Execute   \(F8\)",False) 
Wait(2)
Call TakeScreenShot()

'Verify the grid content
Call VerifyGridCellContentbyID("cntlGRID1",1,"Delivery",0,DT_VL09_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VBELN)

'Select the Row
Call SelectRowGuiGrid("",0,"Delivery",DT_VL09_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VBELN,False)

'Click on Execute
Call ClickButton("Reverse Goods Movement   \(F5\)",False) 
Wait(2)
Call TakeScreenShot()

'Click on continue button in Pop Up
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(5)
Call TakeScreenShot()

'Verify the Labels
VerifyifGuiLabelExists("Goods issue for delivery "&DT_VL09_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VBELN&" canceled")
VerifyifGuiLabelExists("No problems have been logged")

'Click on Continue button
Call ClickButtonIfExist("Continue   \(Enter\)",False)
Wait(5)
Call TakeScreenShot()

'Click on Back
Call ClickButtonIfExist("Back   \(F3\)",False)
Wait(2)

'Click on Back
Call ClickButtonIfExist("Back   \(F3\)",False)
Wait(2)

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

