
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Direct Re- Export without border crossing_p5
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
'.................Test Script Name :Test_Direct Re- Export without border crossing_p5
'.................Author : TCS 	   :Raushan
'................ Creation Date    :20th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Direct Re- Export without border crossing_p5"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Direct Re- Export without border crossing_p5.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode VF03----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter() 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'Display the Invoice Details
Call SetTextbox("Billing document","VBRK-VBELN","",DT_DOC,False)
Call TakeScreenShot()

'Click on Accounting overview
Call ClickButton("Accounting overview   \(F6\)",False)
Wait(1)
Call TakeScreenShot()

'Call GetGridContentByTitle("Documents in Accounting",0,"Doc. Number",1,"DT_ACCOUNTING_DOC_NO")
Call GetGridContentByTitle("Documents in Accounting",0,"Document Number",1,"DT_ACCOUNTING_DOC_NO")
Call TakeScreenShot()


'Call  DoubleClickGuiGridCell("Documents in Accounting",0,1,"Doc. Number",True)
Call  DoubleClickGuiGridCell("Documents in Accounting",0,1,"Document Number",True)

'Click on Back
Call ClickButtonIfExist("Back   \(F3\)",False)
Wait(2)

'Cancel the Pop Up
Call ClickButton("Cancel   \(F12\)",True)
Wait(1)

'Navigate to Billing document;Issue Output To
Call SelectMenuBar("Billing document;Issue Output To")
Call TakeScreenShot()

'Select Row
Call SelectRowGuiTable("SAPLVMSGTABCONTROL","Message type","ZGRS",True)
Call TakeScreenShot()

'Click on Print
Call ClickButton("Print preview   \(Ctrl\+Shift\+F1\)",True)
Wait(1)
Call TakeScreenShot()

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

