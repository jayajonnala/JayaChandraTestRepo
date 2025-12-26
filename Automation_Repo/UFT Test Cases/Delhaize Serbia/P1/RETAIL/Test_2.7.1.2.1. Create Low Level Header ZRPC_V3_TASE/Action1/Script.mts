
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_2.7.1.2.1. Create Low Level Header ZRPC_V3
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
'.................Test Script Name :Test_2.7.1.2.1. Create Low Level Header ZRPC_V3
'.................Author : TCS 	   :Raushan
'................ Creation Date    :16th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_2.7.1.2.1. Create Low Level Header ZRPC_V3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_2.7.1.2.1. Create Low Level Header ZRPC_V3.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode CL24N----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
Call TakeScreenShot()

'Enter the Class Details
Call SetTextbox("Class","RMCLF-CLASN","",DT_CL24N_1110_CLASS,FALSE)
Call SetTextbox("Class Type","RMCLF-KLART","",DT_CL24N_1110_CLASS_TYPE,FALSE)
Call TakeScreenShot()
Call PressEnter() 
Wait(1)
Call PressEnter() 
Call ClickButtonIfExist("Continue   \(Enter\)",True)
wait(2)


'Click on Create new assignments
Call ClickButton("Create new assignments   \(Ctrl\+F9\)",False) 
Wait(1)

Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{LEFT}",1
Wait(2)
wsh.SendKeys "{DOWN}",1
Set wsh= nothing

'Select the Radio button
Call SelectRadioButtonByIndexIfPopupExists("RMCLF-RADIO",1)
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)
wait(2)

'''Call ClickButton("Cancel   \(F12\)",False) 
'''Wait(1)
'''
''''Click on Create new assignments
'''Call ClickButton("Create new assignments   \(Ctrl\+F9\)",False) 
'''Wait(1)
'''
''''Select the Radio button
'''Call SelectRadioButtonByIndexIfPopupExists("RMCLF-RADIO",1)
'''wait(2)
'''Call TakeScreenShot()
'''Call ClickButtonIfExist("Continue   \(Enter\)",True)
'''wait(2)
'''

'Set the Table Data
Call SetTableDataNoRef("SAPLCBCMTC_OBJ_CLASS","CustomerNoSite",1,DT_CL24N_0227_TABLECELL_CUSTOMERNOSITE_0,False)
Wait(2)
Call SetTableDataNoRef("SAPLCBCMTC_OBJ_CLASS","CustomerNoSite",2,DT_CL24N_0227_TABLECELL_CUSTOMERNOSITE_1,False)
Call TakeScreenShot()
Call PressEnter() 
Wait(2)
Call TakeScreenShot()


'Save  the Details
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  
VerifyStatusBar(DT_CL24N_1110_CHECK_TEXT_OF_STATUSBAR)

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

