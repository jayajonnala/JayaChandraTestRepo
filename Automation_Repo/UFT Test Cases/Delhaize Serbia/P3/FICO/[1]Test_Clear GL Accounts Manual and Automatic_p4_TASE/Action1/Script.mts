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


'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Clear AP Accounts - Manual_p4_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 8th April
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Clear AP Accounts - Manual_p4_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Clear AP Accounts - Manual_p4_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
'''----------------------Tcode F-44----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SelectMenuBar("Document;Display")
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Number","RF05L-BELNR","",DT_F44_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_F44_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_F44_0100_FISCAL_YEAR,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

'verify the grid coponents
call VerifyGridCellContent("",1,"Posting Key","",DT_F44_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
call VerifyGridCellContent("",1,"Account","",DT_F44_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
call VerifyGridCellContent("",1,"Amount","",DT_F44_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)

call VerifyGridCellContent("",2,"Posting Key","",DT_F44_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
call VerifyGridCellContent("",2,"Account","",DT_F44_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
call VerifyGridCellContent("",2,"Amount","",DT_F44_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)

Call ClickButton("Display Document Header   \(F5\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyTextBoxContent("Document type","BKPF-BLART","",DT_F44_1710_CHECK_TEXT_OF_DOCUMENT_TYPE,True)
Call VerifyTextBoxContent("Document type","T003T-LTEXT","",Lcase(DT_F44_1710_CHECK_TEXT_OF_DOCUMENT_TYPE_OCC1),True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Continue/Confirm   \(Enter\)",True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

