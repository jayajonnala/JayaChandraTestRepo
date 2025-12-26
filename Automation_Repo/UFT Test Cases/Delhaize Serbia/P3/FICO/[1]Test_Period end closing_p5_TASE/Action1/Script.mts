
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Period end closing_p5
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
'.................Test Script Name :Test_Period end closing_p5
'.................Author : TCS 	   :Raushan
'................ Creation Date    :16th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Period end closing_p5"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_2.7.1.2.1. Create Low Level Header ZRPC.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode KSV3----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

'Enter the details
Call SetTextbox("Cycle","RKAL1-KSCYC","",DT_KSV3_0103_CYCLE,FALSE)
Call SetTextbox("Start Date","T811C-SDATE","",ConvertDate(DT_KSV3_0103_START_DATE),FALSE)
Call TakeScreenShot()
Call PressEnter()     ' 

Call TakeScreenShot()

'Click on First segment
Call ClickButton("First segment   \(F6\)",False)
Call TakeScreenShot()

'Navigate to Senders/Receivers Tab
Call SelectTab("SEG_TABSTRIP","Senders/Receivers",False)
Wait(1)
Call TakeScreenShot()

'Verify the Content
Call VerifyTextBoxContent("Cost Center","KGALK-VALMIN",0,DT_KSV3_0306_CHECK_TEXT_OF_COST_CENTER,False)
Call VerifyTextBoxContent("Cost Center","KGALK-VALMIN",1,DT_KSV3_0306_CHECK_TEXT_OF_COST_CENTER_OCC1,False)

Call FocusTextBox("Cost Element","KGALK-SETID",False)
Call DoubleClick()


Call ClickLabel(DT_KSV3_0200_CHECK_TEXT_OF_FROM_VALUE,0,True)
Set objWsh = CreateObject("WScript.Shell") 
objWsh.SendKeys "{F2}" 
Set objWsh=nothing


Call VerifyTextBoxContent("From value","RGSMV-FROM",0,DT_KSV3_0200_CHECK_TEXT_OF_FROM_VALUE,True)

'Click on Cancel
Call ClickButton("Cancel   \(F12\)",True)

'Click on Cancel
Call ClickButton("Cancel   \(F12\)",True)

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

