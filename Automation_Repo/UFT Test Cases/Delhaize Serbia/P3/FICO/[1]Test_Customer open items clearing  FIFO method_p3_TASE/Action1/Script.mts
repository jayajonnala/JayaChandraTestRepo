'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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

'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Customer open items clearing FIFO method_p3_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 3rd May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Customer open items clearing FIFO method_p3_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Customer open items clearing  FIFO method_p3_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
'''----------------------Tcode ZFIAR_CLEAR_CUST_OI----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Customer","S_KUNNR-LOW","",DT_ZFIAR_CLEAR_CUST_OI_1000_CUSTOMER,False)
Call SetTextbox("Company Code","P_BUKRS","",DT_ZFIAR_CLEAR_CUST_OI_1000_COMPANY_CODE,False)
Call SetTextbox("Clearing date","P_CDATE","",ConvertDateFormat(DT_ZFIAR_CLEAR_CUST_OI_1000_CLEARING_DATE),False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectMenuBar("Program;Execute in Background")
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Output Device","PRI_PARAMS-PDEST","",DT_ZFIAR_CLEAR_CUST_OI_0100_OUTPUT_DEVICE,True)
'Capture the screenshot
Call TakeScreenShot()

Call FocusTextBox("Output Device","PRI_PARAMS-PDEST",True)
Call ClickButton("Continue   \(Shift\+F1\)",True)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Immediate",True)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Save   \(Ctrl\+S\)",True)
Wait(4)
'Capture the screenshot
Call TakeScreenShot()
'Validate If batch is scheduled
Call GetTextStatusBar("DT_ZFIAR_CLEAR_CUST_OI_1000_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_ZFIAR_CLEAR_CUST_OI_1000_CHECK_TEXT_OF_STATUSBAR_OCC1)

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

