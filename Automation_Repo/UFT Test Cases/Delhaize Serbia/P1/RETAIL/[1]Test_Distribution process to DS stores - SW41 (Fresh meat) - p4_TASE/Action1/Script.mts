
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
'................ Creation Date    :16th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Distribution process to DS stores - SW41 (Fresh meat) - p4"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Distribution process to DS stores - SW41 (Fresh meat) - p1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode VL09N----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'Enter the Outbound Delivery No and Radio button
Call SetTextbox("Inbound / Outbound Delivery","I_VBELN-LOW","",DT_VL09_1000_INBOUND__OUTBOUND_DELIVERY,FALSE)
Call SelectRadioButton("LF_ANAUS","Inbound Delivs & Outbound Delivs",False)
Call TakeScreenShot()

'Click on Execute
Call ClickButton("Execute   \(F8\)",False) 
Wait(2)

' VerifyGridCellContent(gridTitle, gridRowNumber, gridColumnName, gridIndex, expectedValue)
Call VerifyGridCellContent("",1,"Delivery","",DT_VL09_1000_INBOUND__OUTBOUND_DELIVERY)
'Verify the grid content
'Call VerifyGridCellContentbyID("cntlGRID1",1,"Delivery",0,DT_VL09_1000_INBOUND__OUTBOUND_DELIVERY)


'Select the Row
Call SelectRowGuiGrid("",0,"Delivery",DT_VL09_1000_INBOUND__OUTBOUND_DELIVERY,False)

'Click on Execute
Call ClickButton("Reverse Goods Movement   \(F5\)",False) 
Wait(2)
Call TakeScreenShot()

'Click on continue button in Pop Up
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(10)

'Verify the Labels
VerifyifGuiLabelExists(DT_VL09_0120_CHECK_TEXT_OF_GOODS_ISSUE_FOR_DELIVERY_CANCELED)
VerifyifGuiLabelExists(DT_VL09_0120_CHECK_TEXT_OF_NO_PROBLEMS_HAVE_BEEN_LOGGED)

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

