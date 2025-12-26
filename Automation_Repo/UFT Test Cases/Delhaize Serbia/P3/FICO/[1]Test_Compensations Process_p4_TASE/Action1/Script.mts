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
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Compensations Process_p4_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 24th March
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Compensations Process_p4_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Compensations Process_p4_TASE.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'''Login'''
'DataRowSet=2
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
'''----------------------Tcode F-30----------------------------
'

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()


Call SelectMenuBar("Document;Display")
Wait(1)


Call VerifyStatusBarMessageType("E")
Call TakeScreenShot()

Call SetTextbox("Document Number","RF05L-BELNR","",DT_F30_0100_DOCUMENT_NUMBER,False)
'Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_ABUMN_0100_FISCAL_YEAR,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_F30_0100_COMPANY_CODE,False)
Call TakeScreenShot()

Call PressEnter()
Call TakeScreenShot()

Call PressEnter()
Call TakeScreenShot()

Call ClickButton("Display Document Header   \(F5\)",False)
Wait(1)
Call TakeScreenShot()

Call VerifyTextBoxContent("Document type","T003T-LTEXT","",Lcase(DT_F30_1710_CHECK_TEXT_OF_DOCUMENT_TYPE),True)

Call ClickButton("Continue/Confirm   \(Enter\)",True)
Call TakeScreenShot()


call VerifyGridCellContent("",1,"Posting Key","",DT_F30_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
call VerifyGridCellContent("",2,"Posting Key","",DT_F30_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

call VerifyGridCellContent("",1,"Account","",DT_F30_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
call VerifyGridCellContent("",2,"Account","",DT_F30_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)


Call LogOff()
Call FinalStatus ()


