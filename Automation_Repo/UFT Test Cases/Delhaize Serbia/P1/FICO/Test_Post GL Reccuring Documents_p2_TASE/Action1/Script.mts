
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Post GL Reccuring Documents_p2
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
'.................Test Script Name : Test_Post GL Reccuring Documents_p2
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 19th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Post GL Reccuring Documents_p2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Post GL Reccuring Documents_p2.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System

'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode SM35----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


'Enter the details
Call SetTextbox("Sess\.:","D0100-MAPN","",DT_SM35_1005_SESS,False)   
Call SetTextbox("Created by:","D0100-CREATOR","",DT_SM35_1005_CREATED_BY,False)
Call TakeScreenShot()
Call PressEnter() 
Call TakeScreenShot()


'Select the Row
Call SelectRowGuiTable("SAPMSBDC_CCTC_APQI","Session Name",DT_SM35_1005_SESS,False)
Call TakeScreenShot()

'Click on Process
Call ClickButton("Process session   \(F8\)",False) 
Wait(2)
Call TakeScreenShot()

'Select Display Error only radio button
Call SelectRadioButtonIfPopupExists("D0300-ERROR","Display errors only")
Call SelectCheckbox("D0300-LOGALL",0,DT_SM35_0300_EXTENDED_LOG,True)
Call SelectCheckbox("D0300-EXPERT",0,DT_SM35_0300_EXPERT_MODE,True)
Call TakeScreenShot()

'Click on Process
Call ClickButton("Process   \(Enter\)",True) 
Wait(2)
Call TakeScreenShot()

'Verify the Message
Call VerifyTextBoxContent("Information Message","MESSTXT1",0,UCASE(DT_SM35_0010_CHECK_TEXT_OF_MESSTXT1),True)

'Click on Session Overview Button
Call ClickButton("Go back to batch input session overview   \(Enter\)",True) 
Wait(2)
Call TakeScreenShot()

'Verify the Status 
Call VerifyTableCellContent(1,"Status","SAPMSBDC_CCTC_APQI",DT_SM35_1010_CHECK_ICONNAME_OF_TABLECELL_STATUS_0)


'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

