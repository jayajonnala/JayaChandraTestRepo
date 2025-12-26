
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_113_Update User Responsible on Cost Center
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
''''''------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

gstrTestCaseName = "Test_113_Update User Responsible on Cost Center"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_MD_ABI067_001 Create ZCXT Retail Customer Local or Foreign_TASE.xls"

''''''------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


''--------TransactionCode-KS02----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter()

'3rd Set Controlling Area - Popup textbox
Call SetTextbox("Controlling Area","SVALD-VALUE","",DT_KS02_0300_CONTROLLING_AREA,True)
Call PressEnter()

'4th Set Cost Center textbox
Call FocusTextBox("Cost Center","CSKSZ-KOSTL",False)     ' - Line (9)
Call SetTextbox("Cost Center","CSKSZ-KOSTL","",DT_KS02_0200_COST_CENTER,False)
Call PressEnter()

Call FocusTextBox("User Responsible","CSKSZ-VERAK_USER",False)
Call VerifyTextBoxContent("User Responsible","CSKSZ-VERAK_USER","",DT_KS02_0300_CHECK_TEXT_OF_USER_RESPONSIBLE,False)

Call SetTextbox("User Responsible","CSKSZ-VERAK_USER","",DT_KS02_0300_USER_RESPONSIBLE,False)
Call PressEnter()

Call FocusTextBox("User Responsible","CSKSZ-VERAK_USER",False)
Call VerifyTextBoxContent("User Responsible","CSKSZ-VERAK_USER","",DT_KS02_0300_USER_RESPONSIBLE,False)
Call SelectMenuBar("Cost Center;Save")

''Verify the Status Bar
Call VerifyStatusBar(DT_KS02_0200_CHECK_TEXT_OF_STATUSBAR)

Call LogOff()
Call FinalStatus ()
